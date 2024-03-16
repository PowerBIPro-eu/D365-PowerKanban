import * as WebApiClient from "xrm-webapi-client";
import { Notification } from "../domain/Notification";
import { ActionDispatch, ActionStateProps } from "./ActionState";
import { AppStateDispatch, AppStateProps } from "./AppState";
import { BoardLane } from "./BoardLane";
import { BoardViewConfig } from "./BoardViewConfig";
import { CardForm, CardSegment } from "./CardForm";
import { ConfigStateProps } from "./ConfigState";
import { Attribute, Metadata } from "./Metadata";

const maxInFilterSize = 500;

const getFieldsFromSegment = (segment: CardSegment): Array<string> =>
    segment.rows.reduce(
        (all, curr) => [...all, ...curr.cells.map((c) => c.field)],
        []
    );

const removeChildren = (parent: Element, childTag: string) => {
    const children = parent.getElementsByTagName(childTag);

    for (let index = children.length - 1; index >= 0; index--) {
        children[index].parentNode.removeChild(children[index]);
    }
};

interface FetchDataOptions {
    additionalFields?: Array<string>;
    hideEmptyLanes?: boolean;
    additionalCondition?: {
        attribute: string;
        operator: string;
        values?: Array<string>;
    };
}

/**
 * <summary>Prepares fetch XML for data retrieval. For secondary data, value conditions are used. If they overflow a certain amount of value tags, we create multiple fetches for gathering all data</summary>
 */
const prepareFetch = (
    fetchXml: string,
    swimLaneSource: string,
    form: CardForm,
    metadata: Metadata,
    options?: FetchDataOptions,
    tempFetches: Array<string> = []
): Array<string> => {
    const formFields = Array.from(
        new Set([
            ...getFieldsFromSegment(form.parsed.header),
            ...getFieldsFromSegment(form.parsed.body),
            ...getFieldsFromSegment(form.parsed.footer),
        ])
    );

    if (formFields.every((f) => f !== swimLaneSource)) {
        formFields.push(swimLaneSource);
    }

    if (formFields.every((f) => f !== metadata.PrimaryNameAttribute)) {
        formFields.push(metadata.PrimaryNameAttribute);
    }

    const ownerField = metadata.Attributes.find(
        (a) => a.LogicalName?.toLowerCase() === "ownerid"
    );

    if (ownerField && formFields.every((f) => f !== ownerField.LogicalName)) {
        formFields.push(ownerField.LogicalName);
    }

    if (options?.additionalFields) {
        options.additionalFields.forEach(() => {
            if (formFields.every((f) => f !== swimLaneSource)) {
                formFields.push(swimLaneSource);
            }
        });
    }

    const safeFetchXml = fetchXml
        ? fetchXml
        : `<fetch no-lock="true"><entity name="${metadata.LogicalName}"></entity></fetch>`;

    const parser = new DOMParser();
    const xml = parser.parseFromString(safeFetchXml, "application/xml");
    const root = xml.getElementsByTagName("fetch")[0];
    const entity = root.getElementsByTagName("entity")[0];
    root.setAttribute("no-lock", "true");

    removeChildren(entity, "attribute");

    Array.from(entity.getElementsByTagName("link-entity")).forEach((l) => {
        removeChildren(l, "attribute");
    });

    [metadata.PrimaryIdAttribute]
        .concat(formFields)
        .concat(options?.additionalFields ?? [])
        .map((a) => {
            const e = xml.createElement("attribute");
            e.setAttribute("name", a);

            return e;
        })
        .forEach((e) => {
            entity.append(e);
        });

    const serializer = new XMLSerializer();

    if (options?.additionalCondition) {
        const filter = xml.createElement("filter");
        let didOverflow = false;
        const c = options?.additionalCondition;

        const condition = xml.createElement("condition");
        condition.setAttribute("attribute", c.attribute);
        condition.setAttribute("operator", c.operator);

        if (c.operator.toLowerCase() === "in") {
            c.values.forEach((v, i) => {
                if (i > maxInFilterSize - 1) {
                    didOverflow = true;
                    return;
                }

                const value = xml.createElement("value");
                value.textContent = v;

                condition.append(value);
            });
        } else if (c.values?.length) {
            condition.setAttribute("value", c.values[0]);
        }

        filter.append(condition);
        entity.append(filter);

        if (didOverflow) {
            return prepareFetch(
                safeFetchXml,
                swimLaneSource,
                form,
                metadata,
                {
                    ...options,
                    additionalCondition: {
                        ...c,
                        values: c.values.slice(maxInFilterSize),
                    },
                },
                [...tempFetches, serializer.serializeToString(xml)]
            );
        }
    }

    const fetch = serializer.serializeToString(xml);

    return [...tempFetches, fetch];
};

export const fetchData = async (
    entityName: string,
    fetchXml: string,
    swimLaneSource: string,
    form: CardForm,
    metadata: Metadata,
    attribute: Attribute,
    isPrimary: boolean,
    appState: AppStateProps,
    options?: FetchDataOptions
): Promise<Array<BoardLane>> => {
    try {
        if (!form) {
            return [];
        }

        const fetches = prepareFetch(
            fetchXml,
            swimLaneSource,
            form,
            metadata,
            !isPrimary || appState.primaryDataIds == null
                ? options
                : {
                      ...options,
                      additionalCondition: {
                          attribute: metadata.PrimaryIdAttribute,
                          operator: "in",
                          values: appState.primaryDataIds.length
                              ? appState.primaryDataIds
                              : ["00000000-0000-0000-0000-000000000000"],
                      },
                  }
        );

        const data: Array<any> = [];

        for (let i = 0; i < fetches.length; i++) {
            const fetch = fetches[i];
            const { value: tempData }: { value: Array<any> } =
                await WebApiClient.Retrieve({
                    entityName: entityName,
                    fetchXml: fetch,
                    returnAllPages: true,
                    headers: [
                        {
                            key: "Prefer",
                            value: 'odata.include-annotations="*"',
                        },
                    ],
                });

            data.push(...tempData);
        }

        const sortedData =
            !isPrimary || appState.primaryDataIds == null
                ? data
                : appState.primaryDataIds
                      .map((id) =>
                          data.find(
                              (d) => d[metadata.PrimaryIdAttribute] === id
                          )
                      )
                      .filter((d) => !!d);

        const lanes =
            attribute.AttributeType === "Boolean"
                ? [
                      attribute.OptionSet.FalseOption,
                      attribute.OptionSet.TrueOption,
                  ]
                : attribute.OptionSet.Options.sort((a, b) => a.State - b.State);

        return sortedData.reduce(
            (all: Array<BoardLane>, record) => {
                const laneSource = record[swimLaneSource];

                if (laneSource == null) {
                    const undefinedLane = all.find(
                        (l) => l.option.Value === null
                    );

                    if (undefinedLane) {
                        undefinedLane.data.push(record);
                    } else {
                        all = [
                            {
                                option: {
                                    Value: null,
                                    Color: "#777",
                                    Label: {
                                        UserLocalizedLabel: { Label: "None" },
                                    },
                                } as any,
                                data: [record],
                            },
                            ...all,
                        ];
                    }

                    return all;
                }
                if (attribute.AttributeType === "Boolean") {
                    const lane = all.find(
                        (l) => l.option && l.option.Value == laneSource
                    );

                    if (lane) {
                        lane.data.push(record);
                    } else {
                        all.push({
                            option: !laneSource ? lanes[0] : lanes[1],
                            data: [record],
                        });
                    }

                    return all;
                }

                const lane = all.find(
                    (l) => l.option && l.option.Value === laneSource
                );

                if (lane) {
                    lane.data.push(record);
                } else {
                    const existingLane = lanes.find(
                        (l) => l.Value === laneSource
                    );

                    if (existingLane) {
                        all.push({ option: existingLane, data: [record] });
                    } else {
                        console.warn(
                            `Found data with non valid option set data, did you reorganize or delete option set values? Data needs to be reorganized then. Value found: ${laneSource}`
                        );
                    }
                }

                return all;
            },
            options?.hideEmptyLanes
                ? []
                : (lanes.map((l) => ({
                      option: l,
                      data: [],
                  })) as Array<BoardLane>)
        );
    } catch (e) {}
};

const groupDataByProperty = (
    primaryLookup: string,
    secondaryLookup: string,
    data: Array<any>
) => {
    const primaryNotificationLookup = `_${primaryLookup}_value`;
    const secondaryNotificationLookup = `_${secondaryLookup}_value`;

    return data.reduce((all, cur) => {
        const id =
            cur[primaryNotificationLookup] ?? cur[secondaryNotificationLookup];

        if (all[id]) {
            all[id].push(cur);
        } else {
            all[id] = [cur];
        }

        return all;
    }, {} as { [key: string]: Array<any> });
};

export const fetchSubscriptions = async (config: BoardViewConfig) => {
    const { value: data }: { value: Array<any> } = await WebApiClient.Retrieve({
        entityName: "oss_subscription",
        queryParams: `?$filter=_ownerid_value eq ${Xrm.Page.context
            .getUserId()
            .replace("{", "")
            .replace("}", "")}&$orderby=createdon desc`,
        returnAllPages: true,
    });

    return groupDataByProperty(
        config.primaryEntity.subscriptionLookup,
        config.secondaryEntity ? config.secondaryEntity.subscriptionLookup : "",
        data
    );
};

export const fetchNotifications = async (
    config: BoardViewConfig
): Promise<{ [key: string]: Array<Notification> }> => {
    const { value: data } = await WebApiClient.Retrieve({
        entityName: "oss_notification",
        queryParams: `?$filter=_ownerid_value eq ${Xrm.Page.context
            .getUserId()
            .replace("{", "")
            .replace("}", "")}&$orderby=createdon desc`,
        returnAllPages: true,
        headers: [{ key: "Prefer", value: 'odata.include-annotations="*"' }],
    });

    const notifications: Array<Notification> = data.map((d: Notification) => ({
        ...d,
        parsed: d.oss_data ? JSON.parse(d.oss_data) : undefined,
    }));

    return groupDataByProperty(
        config.primaryEntity.notificationLookup,
        config.secondaryEntity ? config.secondaryEntity.notificationLookup : "",
        notifications
    );
};

export const refresh = async (
    appDispatch: AppStateDispatch,
    appState: AppStateProps,
    configState: ConfigStateProps,
    actionDispatch: ActionDispatch,
    actionState: ActionStateProps,
    fetchXml?: string,
    selectedForm?: CardForm,
    secondaryFetchXml?: string,
    secondarySelectedForm?: CardForm
) => {
    actionDispatch({ type: "setWorkIndicator", payload: true });

    try {
        const data = await fetchData(
            configState.config.primaryEntity.logicalName,
            fetchXml,
            configState.config.primaryEntity.swimLaneSource,
            selectedForm ?? actionState.selectedForm,
            configState.metadata,
            configState.separatorMetadata,
            true,
            appState
        );
        appDispatch({ type: "setBoardData", payload: data });

        if (configState.config.secondaryEntity) {
            const secondaryData = await fetchData(
                configState.config.secondaryEntity.logicalName,
                secondaryFetchXml ?? actionState.selectedSecondaryView.fetchxml,
                configState.config.secondaryEntity.swimLaneSource,
                secondarySelectedForm ?? actionState.selectedSecondaryForm,
                configState.secondaryMetadata[
                    configState.config.secondaryEntity.logicalName
                ],
                configState.secondarySeparatorMetadata,
                false,
                appState,
                {
                    additionalFields: [
                        configState.config.secondaryEntity.parentLookup,
                    ],
                    additionalCondition: {
                        attribute:
                            configState.config.secondaryEntity.parentLookup,
                        operator: "in",
                        values: data.some((d) => d.data.length > 0)
                            ? data.reduce(
                                  (all, d) => [
                                      ...all,
                                      ...d.data.map(
                                          (laneData) =>
                                              laneData[
                                                  configState.metadata
                                                      .PrimaryIdAttribute
                                              ] as string
                                      ),
                                  ],
                                  [] as Array<string>
                              )
                            : ["00000000-0000-0000-0000-000000000000"],
                    },
                }
            );

            appDispatch({ type: "setSecondaryData", payload: secondaryData });
        }

        const notifications = await fetchNotifications(configState.config);
        appDispatch({ type: "setNotifications", payload: notifications });
    } catch (e) {}

    actionDispatch({ type: "setWorkIndicator", payload: false });
};

export const extractTextFromAttribute = (
    data: { [key: string]: any },
    displayField: string
) => {
    if (!data) {
        return "";
    }

    return (
        data[`${displayField}@OData.Community.Display.V1.FormattedValue`] ??
        data[
            `_${displayField}_value@OData.Community.Display.V1.FormattedValue`
        ] ??
        data[displayField]?.toString() ??
        ""
    );
};

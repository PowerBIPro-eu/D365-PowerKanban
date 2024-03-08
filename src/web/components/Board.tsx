import { ICardStyles } from "@uifabric/react-cards";
import * as React from "react";
import * as WebApiClient from "xrm-webapi-client";
import { useActionContext } from "../domain/ActionState";
import { useAppContext } from "../domain/AppState";
import { BoardLane } from "../domain/BoardLane";
import { BoardEntity, BoardViewConfig } from "../domain/BoardViewConfig";
import { CardForm, parseCardForm } from "../domain/CardForm";
import { useConfigContext } from "../domain/ConfigState";
import { formatGuid } from "../domain/GuidFormatter";
import {
    loadExternalResource,
    loadExternalScript,
} from "../domain/LoadExternalResource";
import { Attribute, Metadata, Option } from "../domain/Metadata";
import { RecordFilter } from "../domain/RecordFilter";
import { SavedQuery } from "../domain/SavedQuery";
import {
    fetchData,
    fetchNotifications,
    fetchSubscriptions,
    refresh,
} from "../domain/fetchData";
import { DndContainer } from "./DndContainer";
import { Lane } from "./Lane";
import { Tile } from "./Tile";

const determineAttributeUrl = (attribute: Attribute) => {
    if (attribute.AttributeType === "Picklist") {
        return "Microsoft.Dynamics.CRM.PicklistAttributeMetadata";
    }

    if (attribute.AttributeType === "Status") {
        return "Microsoft.Dynamics.CRM.StatusAttributeMetadata";
    }

    if (attribute.AttributeType === "State") {
        return "Microsoft.Dynamics.CRM.StateAttributeMetadata";
    }

    if (attribute.AttributeType === "Boolean") {
        return "Microsoft.Dynamics.CRM.BooleanAttributeMetadata";
    }

    throw new Error(
        `Type ${attribute.AttributeType} is not allowed as swim lane separator.`
    );
};

export type DisplayState = "simple" | "advanced";

export const Board = () => {
    const [appState, appDispatch] = useAppContext();
    const [actionState, actionDispatch] = useActionContext();
    const [configState, configDispatch] = useConfigContext();

    const [secondaryViews, setSecondaryViews] = React.useState<
        Array<SavedQuery>
    >([]);
    const [cardForms, setCardForms] = React.useState<Array<CardForm>>([]);
    const [secondaryCardForms, setSecondaryCardForms] = React.useState<
        Array<CardForm>
    >([]);
    const [stateFilters, setStateFilters] = React.useState<Array<Option>>([]);
    const [secondaryStateFilters, setSecondaryStateFilters] = React.useState<
        Array<Option>
    >([]);
    const [displayState, setDisplayState] = React.useState<DisplayState>(
        "simple" as any
    );
    const [appliedSearchText, setAppliedSearch] = React.useState(undefined);
    const [showNotificationRecordsOnly, setShowNotificationRecordsOnly] =
        React.useState(false);
    const [error, setError] = React.useState(undefined);
    const [customStyle, setCustomStyle] = React.useState(undefined);

    const [primaryFilters, setPrimaryFilters] = React.useState(
        [] as Array<RecordFilter>
    );
    const [secondaryFilters, setSecondaryFilters] = React.useState(
        [] as Array<RecordFilter>
    );

    const isFirstRun = React.useRef(true);

    if (error) {
        throw error;
    }

    React.useEffect(() => {
        if (!actionState.selectedRecords) {
            return;
        }

        const selectedRecords = Object.keys(actionState.selectedRecords).reduce(
            (all, cur) =>
                actionState.selectedRecords[cur] ? [...all, cur] : all,
            []
        );

        if (selectedRecords.length) {
            appState.pcfContext.parameters.primaryDataSet.setSelectedRecordIds(
                selectedRecords
            );
        } else {
            appState.pcfContext.parameters.primaryDataSet.clearSelectedRecordIds();
        }
    }, [actionState.selectedRecords]);

    const openRecord = React.useCallback((reference: Xrm.LookupValue) => {
        appState.pcfContext.parameters.primaryDataSet.openDatasetItem(
            reference as any
        );
    }, []);

    const getOrSetCachedJsonObjects = async (
        cachedKey: string,
        generator: () => Promise<any>
    ) => {
        const currentCacheKey = `${
            (appState.pcfContext as any).orgSettings.uniqueName
        }_${cachedKey}`;
        const cachedEntry = sessionStorage.getItem(currentCacheKey);

        if (cachedEntry) {
            return JSON.parse(cachedEntry);
        }

        const entry = await Promise.resolve(generator());
        sessionStorage.setItem(currentCacheKey, JSON.stringify(entry));

        return entry;
    };

    const getOrSetJsonObject = async (
        cacheKey: string,
        generator: () => Promise<any>
    ) => {
        if (configState.config.cachingEnabled) {
            return getOrSetCachedJsonObjects(cacheKey, generator);
        }

        return generator();
    };

    const fetchSeparatorMetadata = async (
        entity: string,
        swimLaneSource: string,
        metadata: Metadata
    ) => {
        const cacheKey = `__d365powerkanban_entity_${entity}_field_${swimLaneSource}`;
        const generator = async () => {
            const field = metadata.Attributes.find(
                (a) =>
                    a.LogicalName.toLowerCase() === swimLaneSource.toLowerCase()
            )!;
            const typeUrl = determineAttributeUrl(field);

            const response: Attribute = await WebApiClient.Retrieve({
                entityName: "EntityDefinition",
                queryParams: `(LogicalName='${entity}')/Attributes(LogicalName='${field.LogicalName}')/${typeUrl}?$expand=OptionSet`,
            });
            return response;
        };

        return getOrSetJsonObject(cacheKey, generator);
    };

    const fetchMetadata = async (entity: string) => {
        const cacheKey = `__d365powerkanban_entity_${entity}`;
        const generator = async () => {
            const response = await WebApiClient.Retrieve({
                entityName: "EntityDefinition",
                queryParams: `(LogicalName='${entity}')?$expand=Attributes`,
            });
            return response;
        };

        return getOrSetJsonObject(cacheKey, generator);
    };

    const fetchViews = async (entity: string) => {
        const cacheKey = `__d365powerkanban_views_${entity}`;
        const generator = async () => {
            const response = await WebApiClient.Retrieve({
                entityName: "savedquery",
                queryParams: `?$select=layoutxml,fetchxml,savedqueryid,name&$filter=returnedtypecode eq '${entity}' and querytype eq 0 and statecode eq 0&$orderby=name`,
            });
            return response;
        };

        return getOrSetJsonObject(cacheKey, generator);
    };

    const fetchForms = async (entity: string) => {
        const cacheKey = `__d365powerkanban_forms_${entity}`;
        const generator = async () => {
            const response = await WebApiClient.Retrieve({
                entityName: "systemform",
                queryParams: `?$select=formxml,name&$filter=objecttypecode eq '${entity}' and type eq 11`,
            });
            return response;
        };

        return getOrSetJsonObject(cacheKey, generator);
    };

    const fetchConfig = async (configId: string): Promise<BoardViewConfig> => {
        const config = await WebApiClient.Retrieve({
            entityName: "oss_powerkanbanconfig",
            entityId: configId,
            queryParams: "?$select=oss_value",
        });

        return JSON.parse(config.oss_value);
    };

    const getConfigId = async () => {
        if (configState.configId) {
            return configState.configId;
        }

        const userId = formatGuid(Xrm.Page.context.getUserId());
        const user = await WebApiClient.Retrieve({
            entityName: "systemuser",
            entityId: userId,
            queryParams: "?$select=oss_defaultboardid",
        });

        return user.oss_defaultboardid;
    };

    const loadConfig = async () => {
        try {
            appDispatch({ type: "setSecondaryData", payload: [] });
            appDispatch({ type: "setBoardData", payload: [] });
            setCustomStyle(undefined);

            const configId = await getConfigId();

            if (!configId) {
                actionDispatch({
                    type: "setConfigSelectorDisplayState",
                    payload: true,
                });
                return;
            }

            actionDispatch({
                type: "setProgressText",
                payload: "Fetching configuration",
            });
            const config = await fetchConfig(configId);

            if (config.customScriptUrl) {
                actionDispatch({
                    type: "setProgressText",
                    payload: "Loading custom scripts",
                });
                await loadExternalScript(config.customScriptUrl);
            }

            if (config.customStyleUrl) {
                actionDispatch({
                    type: "setProgressText",
                    payload: "Loading custom styles",
                });
                setCustomStyle(
                    await loadExternalResource(config.customStyleUrl)
                );
            }

            if (
                config.defaultDisplayState &&
                (["simple", "advanced"] as Array<DisplayState>).includes(
                    config.defaultDisplayState
                )
            ) {
                setDisplayState(config.defaultDisplayState);
            }

            configDispatch({ type: "setConfig", payload: config });
        } catch (e) {
            actionDispatch({ type: "setProgressText", payload: undefined });
            setError(e);
        }
    };

    const initializeConfig = async () => {
        if (!configState.config) {
            return;
        }

        try {
            actionDispatch({
                type: "setProgressText",
                payload: "Fetching meta data",
            });

            const metadata = await fetchMetadata(
                configState.config.primaryEntity.logicalName
            );
            const attributeMetadata = await fetchSeparatorMetadata(
                configState.config.primaryEntity.logicalName,
                configState.config.primaryEntity.swimLaneSource,
                metadata
            );

            const notificationMetadata = await fetchMetadata(
                "oss_notification"
            );
            configDispatch({
                type: "setSecondaryMetadata",
                payload: {
                    entity: "oss_notification",
                    data: notificationMetadata,
                },
            });

            let secondaryMetadata: Metadata;
            let secondaryAttributeMetadata: Attribute;

            if (configState.config.secondaryEntity) {
                secondaryMetadata = await fetchMetadata(
                    configState.config.secondaryEntity.logicalName
                );
                secondaryAttributeMetadata = await fetchSeparatorMetadata(
                    configState.config.secondaryEntity.logicalName,
                    configState.config.secondaryEntity.swimLaneSource,
                    secondaryMetadata
                );

                configDispatch({
                    type: "setSecondaryMetadata",
                    payload: {
                        entity: configState.config.secondaryEntity.logicalName,
                        data: secondaryMetadata,
                    },
                });
                configDispatch({
                    type: "setSecondarySeparatorMetadata",
                    payload: secondaryAttributeMetadata,
                });
            }

            configDispatch({ type: "setMetadata", payload: metadata });
            configDispatch({
                type: "setSeparatorMetadata",
                payload: attributeMetadata,
            });
            actionDispatch({
                type: "setProgressText",
                payload: "Fetching views",
            });

            let defaultSecondaryView;
            if (configState.config.secondaryEntity) {
                const { value: secondaryViews }: { value: Array<SavedQuery> } =
                    await fetchViews(
                        configState.config.secondaryEntity.logicalName
                    );
                setSecondaryViews(
                    secondaryViews.filter(
                        (v) =>
                            (!configState.config.secondaryEntity.hiddenViews ||
                                !configState.config.secondaryEntity.hiddenViews.some(
                                    (h) =>
                                        v.name?.toLowerCase() ===
                                            h?.toLowerCase() ||
                                        v.savedqueryid?.toLowerCase() ===
                                            h?.toLowerCase()
                                )) &&
                            (!configState.config.secondaryEntity.visibleViews ||
                                configState.config.secondaryEntity.visibleViews.some(
                                    (h) =>
                                        v.name?.toLowerCase() ===
                                            h?.toLowerCase() ||
                                        v.savedqueryid?.toLowerCase() ===
                                            h?.toLowerCase()
                                ))
                    )
                );

                defaultSecondaryView = configState.config.secondaryEntity
                    .defaultView
                    ? secondaryViews.find((v) =>
                          [v.savedqueryid, v.name]
                              .map((i) => i.toLowerCase())
                              .includes(
                                  configState.config.secondaryEntity.defaultView.toLowerCase()
                              )
                      ) ?? secondaryViews[0]
                    : secondaryViews[0];

                actionDispatch({
                    type: "setSelectedSecondaryView",
                    payload: defaultSecondaryView,
                });
            }

            actionDispatch({
                type: "setProgressText",
                payload: "Fetching forms",
            });

            const { value: forms } = await fetchForms(
                configState.config.primaryEntity.logicalName
            );
            const processedForms: Array<CardForm> = forms.map((f: any) => ({
                ...f,
                parsed: parseCardForm(f),
            }));
            processedForms.sort((a, b) => a.parsed.order - b.parsed.order);
            setCardForms(processedForms);

            const { value: notificationForms } = await fetchForms(
                "oss_notification"
            );
            const processedNotificationForms: Array<CardForm> =
                notificationForms.map((f: any) => ({
                    ...f,
                    parsed: parseCardForm(f),
                }));
            processedNotificationForms.sort(
                (a, b) => a.parsed.order - b.parsed.order
            );
            configDispatch({
                type: "setNotificationForm",
                payload: processedNotificationForms[0],
            });

            let defaultSecondaryForm;
            if (configState.config.secondaryEntity) {
                const { value: forms } = await fetchForms(
                    configState.config.secondaryEntity.logicalName
                );
                const processedSecondaryForms: Array<CardForm> = forms.map(
                    (f: any) => ({ ...f, parsed: parseCardForm(f) })
                );
                processedSecondaryForms.sort(
                    (a, b) => a.parsed.order - b.parsed.order
                );
                setSecondaryCardForms(processedSecondaryForms);

                defaultSecondaryForm = processedSecondaryForms[0];
                actionDispatch({
                    type: "setSelectedSecondaryForm",
                    payload: defaultSecondaryForm,
                });
            }

            const defaultForm = processedForms[0];

            if (!defaultForm) {
                actionDispatch({ type: "setProgressText", payload: undefined });
                return Xrm.Utility.alertDialog(
                    `Did not find any card forms for ${configState.config.primaryEntity.logicalName}, please create one.`,
                    () => {}
                );
            }

            actionDispatch({ type: "setSelectedForm", payload: defaultForm });

            actionDispatch({
                type: "setProgressText",
                payload: "Fetching subscriptions",
            });
            const subscriptions = await fetchSubscriptions(configState.config);
            appDispatch({ type: "setSubscriptions", payload: subscriptions });

            actionDispatch({
                type: "setProgressText",
                payload: "Fetching notifications",
            });
            const notifications = await fetchNotifications(configState.config);
            appDispatch({ type: "setNotifications", payload: notifications });

            actionDispatch({
                type: "setProgressText",
                payload: "Fetching data",
            });

            const data = await fetchData(
                configState.config.primaryEntity.logicalName,
                null,
                configState.config.primaryEntity.swimLaneSource,
                defaultForm,
                metadata,
                attributeMetadata,
                true,
                appState,
                {}
            );

            if (configState.config.secondaryEntity) {
                const secondaryData = await fetchData(
                    configState.config.secondaryEntity.logicalName,
                    defaultSecondaryView.fetchxml,
                    configState.config.secondaryEntity.swimLaneSource,
                    defaultSecondaryForm,
                    secondaryMetadata,
                    secondaryAttributeMetadata,
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
                            values: data.some((d) => d.data.length > 1)
                                ? data.reduce(
                                      (all, d) => [
                                          ...all,
                                          ...d.data.map(
                                              (laneData) =>
                                                  laneData[
                                                      metadata
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
                appDispatch({
                    type: "setSecondaryData",
                    payload: secondaryData,
                });
            }

            appDispatch({ type: "setBoardData", payload: data });
            actionDispatch({ type: "setProgressText", payload: undefined });
        } catch (e) {
            actionDispatch({ type: "setProgressText", payload: undefined });
            setError(e);
        }
    };

    React.useEffect(() => {
        loadConfig();
    }, [configState.configId]);

    React.useEffect(() => {
        initializeConfig();
    }, [configState.config]);

    const refreshBoard = async () => {
        appState.pcfContext.parameters.primaryDataSet.refresh();
    };

    React.useEffect(() => {
        if (isFirstRun.current) {
            isFirstRun.current = false;
            return;
        }

        if (!configState || !configState.config) {
            return;
        }

        refresh(
            appDispatch,
            appState,
            configState,
            actionDispatch,
            actionState
        );
    }, [appState.primaryDataIds]);

    const advancedTileStyle = React.useMemo(
        () => ({ margin: "5px" as React.ReactText } as ICardStyles),
        []
    );

    const filterForSearchText = (d: BoardLane) =>
        !appliedSearchText
            ? d
            : {
                  ...d,
                  data: d.data.filter((data) =>
                      Object.keys(data).some((k) =>
                          `${data[k]}`
                              .toLowerCase()
                              .includes(appliedSearchText.toLowerCase())
                      )
                  ),
              };

    const filterForNotifications = (d: BoardLane) =>
        !showNotificationRecordsOnly
            ? d
            : {
                  ...d,
                  data: d.data.filter(
                      (data) =>
                          appState.notifications &&
                          appState.notifications[
                              data[configState.metadata.PrimaryIdAttribute]
                          ] &&
                          appState.notifications[
                              data[configState.metadata.PrimaryIdAttribute]
                          ].length
                  ),
              };

    const filterLanes = (
        d: BoardLane,
        e: BoardEntity,
        filters: Array<Option>
    ) => {
        const isStateVisible =
            !filters.length || filters.some((f) => f.Value === d.option.Value);
        const isVisibleLane =
            !e.visibleLanes || e.visibleLanes.some((l) => l === d.option.Value);
        const isHiddenLane = e.hiddenLanes?.some((l) => l === d.option.Value);

        return isStateVisible && isVisibleLane && !isHiddenLane;
    };

    const filterPrimaryLanes = (d: BoardLane) =>
        filterLanes(d, configState?.config.primaryEntity, stateFilters);
    const filterSecondaryLanes = (d: BoardLane) =>
        filterLanes(
            d,
            configState?.config.secondaryEntity,
            secondaryStateFilters
        );

    const advancedData = React.useMemo(() => {
        return (
            displayState === "advanced" &&
            appState.boardData &&
            appState.boardData
                .filter(filterPrimaryLanes)
                .map(filterForSearchText)
                .map(filterForNotifications)
                .reduce(
                    (all, curr) =>
                        all.concat(
                            curr.data
                                .filter((d) =>
                                    appState.secondaryData.some((t) =>
                                        t.data.some(
                                            (tt) =>
                                                tt[
                                                    `_${configState.config.secondaryEntity.parentLookup}_value`
                                                ] ===
                                                d[
                                                    configState.metadata
                                                        .PrimaryIdAttribute
                                                ]
                                        )
                                    )
                                )
                                .map((d) => {
                                    const secondaryData = appState.secondaryData
                                        .filter(filterSecondaryLanes)
                                        .map((s) => ({
                                            ...s,
                                            data: s.data.filter(
                                                (sd) =>
                                                    sd[
                                                        `_${configState.config.secondaryEntity.parentLookup}_value`
                                                    ] ===
                                                    d[
                                                        configState.metadata
                                                            .PrimaryIdAttribute
                                                    ]
                                            ),
                                        }));

                                    const secondarySubscriptions = Object.keys(
                                        appState.subscriptions
                                    )
                                        .filter((k) =>
                                            secondaryData.some((d) =>
                                                d.data.some(
                                                    (r) =>
                                                        r[
                                                            configState
                                                                .secondaryMetadata[
                                                                configState
                                                                    .config
                                                                    .secondaryEntity
                                                                    .logicalName
                                                            ].PrimaryIdAttribute
                                                        ] === k
                                                )
                                            )
                                        )
                                        .reduce(
                                            (all, cur) => ({
                                                ...all,
                                                [cur]: appState.subscriptions[
                                                    cur
                                                ],
                                            }),
                                            {}
                                        );

                                    const secondaryNotifications = Object.keys(
                                        appState.notifications
                                    )
                                        .filter((k) =>
                                            secondaryData.some((d) =>
                                                d.data.some(
                                                    (r) =>
                                                        r[
                                                            configState
                                                                .secondaryMetadata[
                                                                configState
                                                                    .config
                                                                    .secondaryEntity
                                                                    .logicalName
                                                            ].PrimaryIdAttribute
                                                        ] === k
                                                )
                                            )
                                        )
                                        .reduce(
                                            (all, cur) => ({
                                                ...all,
                                                [cur]: appState.notifications[
                                                    cur
                                                ],
                                            }),
                                            {}
                                        );

                                    return (
                                        <Tile
                                            notifications={
                                                !appState.notifications
                                                    ? []
                                                    : appState.notifications[
                                                          d[
                                                              configState
                                                                  .metadata
                                                                  .PrimaryIdAttribute
                                                          ]
                                                      ] ?? []
                                            }
                                            borderColor={
                                                curr.option.Color ?? "#3b79b7"
                                            }
                                            cardForm={actionState.selectedForm}
                                            metadata={configState.metadata}
                                            key={`tile_${
                                                d[
                                                    configState.metadata
                                                        .PrimaryIdAttribute
                                                ]
                                            }`}
                                            style={advancedTileStyle}
                                            data={d}
                                            refresh={refreshBoard}
                                            searchText={appliedSearchText}
                                            subscriptions={
                                                !appState.subscriptions
                                                    ? []
                                                    : appState.subscriptions[
                                                          d[
                                                              configState
                                                                  .metadata
                                                                  .PrimaryIdAttribute
                                                          ]
                                                      ] ?? []
                                            }
                                            selectedSecondaryForm={
                                                actionState.selectedSecondaryForm
                                            }
                                            secondarySubscriptions={
                                                secondarySubscriptions
                                            }
                                            secondaryNotifications={
                                                secondaryNotifications
                                            }
                                            config={
                                                configState.config.primaryEntity
                                            }
                                            separatorMetadata={
                                                configState.separatorMetadata
                                            }
                                            preventDrag={true}
                                            secondaryData={secondaryData}
                                            openRecord={openRecord}
                                            isSelected={
                                                actionState.selectedRecords &&
                                                actionState.selectedRecords[
                                                    d[
                                                        configState.metadata
                                                            .PrimaryIdAttribute
                                                    ]
                                                ]
                                            }
                                        />
                                    );
                                })
                        ),
                    []
                )
        );
    }, [
        displayState,
        showNotificationRecordsOnly,
        appState.boardData,
        appState.secondaryData,
        stateFilters,
        secondaryStateFilters,
        appliedSearchText,
        appState.notifications,
        appState.subscriptions,
        actionState.selectedSecondaryForm,
        actionState.selectedRecords,
        configState.configId,
    ]);

    const simpleData = React.useMemo(() => {
        return (
            appState.boardData &&
            appState.boardData
                .filter(filterPrimaryLanes)
                .map(filterForSearchText)
                .map(filterForNotifications)
                .map((d) => (
                    <Lane
                        key={`lane_${d.option?.Value ?? "fallback"}`}
                        cardForm={actionState.selectedForm}
                        metadata={configState.metadata}
                        refresh={refreshBoard}
                        subscriptions={appState.subscriptions}
                        searchText={appliedSearchText}
                        config={configState.config.primaryEntity}
                        separatorMetadata={configState.separatorMetadata}
                        openRecord={openRecord}
                        selectedRecords={actionState.selectedRecords}
                        lane={{
                            ...d,
                            data: d.data.filter(
                                (r) =>
                                    displayState === "simple" ||
                                    (appState.secondaryData &&
                                        appState.secondaryData.every((t) =>
                                            t.data.every(
                                                (tt) =>
                                                    tt[
                                                        `_${configState.config.secondaryEntity.parentLookup}_value`
                                                    ] !==
                                                    r[
                                                        configState.metadata
                                                            .PrimaryIdAttribute
                                                    ]
                                            )
                                        ))
                            ),
                        }}
                    />
                ))
        );
    }, [
        displayState,
        showNotificationRecordsOnly,
        appState.boardData,
        appState.subscriptions,
        stateFilters,
        appState.secondaryData,
        appliedSearchText,
        appState.notifications,
        configState.configId,
        actionState.selectedRecords,
    ]);

    React.useEffect(() => {
        if (!actionState.selectedForm?.parsed) {
            setPrimaryFilters([]);
            return;
        }

        const fields = [
            ...actionState.selectedForm.parsed.header.rows.reduce(
                (all, cur) => [...all, ...cur.cells.map((c) => c.field)],
                [] as Array<string>
            ),
            ...actionState.selectedForm.parsed.body.rows.reduce(
                (all, cur) => [...all, ...cur.cells.map((c) => c.field)],
                [] as Array<string>
            ),
            ...actionState.selectedForm.parsed.footer.rows.reduce(
                (all, cur) => [...all, ...cur.cells.map((c) => c.field)],
                [] as Array<string>
            ),
        ];

        setPrimaryFilters(
            fields.map((f) => ({
                logicalName: f,
                displayName: configState.metadata.Attributes.find(
                    (a) => a.LogicalName === f
                )?.DisplayName.UserLocalizedLabel.Label,
                operator: "contains",
            }))
        );
    }, [appState.boardData]);

    React.useEffect(() => {
        if (!actionState.selectedSecondaryForm?.parsed) {
            setSecondaryFilters([]);
            return;
        }

        const fields = [
            ...actionState.selectedSecondaryForm.parsed.header.rows.reduce(
                (all, cur) => [...all, ...cur.cells.map((c) => c.field)],
                [] as Array<string>
            ),
            ...actionState.selectedSecondaryForm.parsed.body.rows.reduce(
                (all, cur) => [...all, ...cur.cells.map((c) => c.field)],
                [] as Array<string>
            ),
            ...actionState.selectedSecondaryForm.parsed.footer.rows.reduce(
                (all, cur) => [...all, ...cur.cells.map((c) => c.field)],
                [] as Array<string>
            ),
        ];

        setSecondaryFilters(
            fields.map((f) => ({
                logicalName: f,
                displayName: configState.secondaryMetadata[
                    configState.config.secondaryEntity.logicalName
                ].Attributes.find((a) => a.LogicalName === f)?.DisplayName
                    .UserLocalizedLabel.Label,
                operator: "contains",
            }))
        );
    }, [appState.secondaryData]);

    return (
        <div
            style={{ height: "100%", display: "flex", flexDirection: "column" }}
        >
            {customStyle && <style>{customStyle}</style>}
            <DndContainer>
                {displayState === "advanced" && (
                    <div
                        id="advancedContainer"
                        style={{
                            display: "flex",
                            flexDirection: "column",
                            overflow: "auto",
                        }}
                    >
                        {advancedData}
                    </div>
                )}
                {displayState === "simple" && (
                    <div
                        id="flexContainer"
                        style={{
                            display: "flex",
                            flexDirection: "row",
                            overflow: "auto",
                            flex: "1",
                        }}
                    >
                        {simpleData}
                    </div>
                )}
            </DndContainer>
        </div>
    );
};

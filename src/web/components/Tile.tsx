import { Badge } from "@fluentui/react-badge";
import { Card as CardFluent, CardHeader } from "@fluentui/react-card";
import {
    Body1Strong,
    Button,
    Field,
    Persona,
    ProgressBar,
} from "@fluentui/react-components";
import { Divider } from "@fluentui/react-divider";
import { FluentProvider } from "@fluentui/react-provider";
import { Caption1, Subtitle2, Text } from "@fluentui/react-text";
import { ICardStyles } from "@uifabric/react-cards";
import * as React from "react";
import { DragSourceMonitor, useDrag } from "react-dnd";
import * as WebApiClient from "xrm-webapi-client";
import { useActionDispatch } from "../domain/ActionState";
import { useAppContext, useAppDispatch } from "../domain/AppState";
import { BoardLane } from "../domain/BoardLane";
import { BoardEntity } from "../domain/BoardViewConfig";
import { CardForm } from "../domain/CardForm";
import { useConfigState } from "../domain/ConfigState";
import { FetchUserAvatar } from "../domain/FetchUserInfo";
import { FlyOutForm } from "../domain/FlyOutForm";
import { ItemTypes } from "../domain/ItemTypes";
import { Attribute, Metadata, Option } from "../domain/Metadata";
import { Notification } from "../domain/Notification";
import { Subscription } from "../domain/Subscription";
import { makeStyles, mergeClasses } from "@fluentui/react-components";
import {
    CalendarCancel16Regular,
    CalendarLtr20Regular,
    GanttChart24Regular,
    Open16Regular,
} from "@fluentui/react-icons";
import { tokens, webLightTheme } from "@fluentui/tokens";
import { shorthands } from "@griffel/core/";
import PlantPackageKam from "./TileContent/PlantPackageKAM";
import CostingSheetContent from "./TileContent/CostingSheetContent";
import PlantPackageCE from "./TileContent/PlantPackageCE";
import PlantPackageLog from "./TileContent/PlantPackageLOG";
import PlantPackagePo from "./TileContent/PlantPackagePO";

interface TileProps {
    borderColor: string;
    cardForm: CardForm;
    config: BoardEntity;
    data: any;
    dndType?: string;
    laneOption?: Option;
    metadata: Metadata;
    notifications?: Array<Notification>;
    searchText: string;
    secondaryData?: Array<BoardLane>;
    secondaryNotifications?: { [key: string]: Array<Notification> };
    secondarySubscriptions?: { [key: string]: Array<Subscription> };
    selectedSecondaryForm?: CardForm;
    separatorMetadata: Attribute;
    style?: ICardStyles;
    subscriptions: Array<Subscription>;
    refresh: () => Promise<void>;
    preventDrag?: boolean;
    openRecord: (reference: Xrm.LookupValue) => void;
    isSelected: boolean;
    isSecondary?: boolean;
}

const TileRender = (props: TileProps) => {
    const [completions, setCompletions] = React.useState<any>({
        bip: "BIPs 3/7",
        cs: "CS 7/7",
        fs: "Feasiblity Study 5/6",
    });
    const [appState, appDispatch] = useAppContext();
    const actionDispatch = useActionDispatch();
    const [, setOverriddenStyle] = React.useState({} as ICardStyles);
    const [personaUrl, setPersonaUrl] = React.useState<string>(undefined);
    const [projectImgUrl, setProjectImageUrl] =
        React.useState<string>(undefined);
    const [assignees, setAssignees] = React.useState<
        Array<{ [key: string]: any }>
    >([]);
    const [attachments, setAttachments] = React.useState<
        Array<{ [key: string]: any }>
    >([]);
    const [managers, setManagers] = React.useState<Array<string>>([]);

    React.useEffect(() => {
        setCompletions({
            bip: "BIPs 3/7",
            cs: "CS 7/7",
            fs: "Feasiblity Study 5/6",
        });
    }, []);

    if (props.config.persona) {
        React.useEffect(() => {
            const personaAttribute = props.metadata.Attributes.find(
                (a) => a.LogicalName?.toLowerCase() === props.config.persona
            );

            if (
                !personaAttribute ||
                personaAttribute.AttributeType !== "Owner"
            ) {
                return;
            }

            const ownerType =
                props.data[
                    `_${props.config.persona}_value@Microsoft.Dynamics.CRM.lookuplogicalname`
                ];
            const ownerId = props.data[`_${props.config.persona}_value`];

            if (ownerType !== "systemuser" || !ownerId) {
                return;
            }

            FetchUserAvatar(ownerId).then((url) => {
                setPersonaUrl(url);
            });
        }, [props.data[props.config.persona]]);
    }

    const context = {
        showForm: (form: FlyOutForm) => {
            return new Promise((resolve, reject) => {
                form.resolve = resolve;
                form.reject = reject;

                actionDispatch({ type: "setFlyOutForm", payload: form });
            });
        },
        refresh: props.refresh,
        setWorkIndicator: (working: boolean) => {
            return actionDispatch({
                type: "setWorkIndicator",
                payload: working,
            });
        },
        data: props.data,
        WebApiClient: WebApiClient,
    };

    const accessFunc = (identifier: string): Function => {
        const path = identifier.split(".");
        return path.reduce(
            (all, cur) => (!all ? undefined : (all as any)[cur]),
            window
        ) as any;
    };

    const [, drag] = useDrag<
        { id: string; sourceLane: Option; type: string } | undefined,
        undefined,
        { isDragging: boolean }
    >({
        item: {
            id: props.data[props.metadata.PrimaryIdAttribute],
            sourceLane: props.laneOption,
            type: props.dndType ?? ItemTypes.Tile,
        } as any,
        end: (
            item: { id: string; sourceLane: Option } | undefined,
            monitor: DragSourceMonitor
        ) => {
            const asyncEnd = async (
                item: { id: string; sourceLane: Option } | undefined,
                monitor: DragSourceMonitor
            ) => {
                const dropResult = monitor.getDropResult();

                if (
                    !dropResult ||
                    dropResult?.option?.Value == null ||
                    dropResult.option.Value === item.sourceLane.Value
                ) {
                    return;
                }

                try {
                    let preventDefault = false;

                    if (props.config.transitionCallback) {
                        const eventContext = {
                            ...context,
                            target: dropResult.option,
                        };

                        const funcRef = accessFunc(
                            props.config.transitionCallback
                        ) as any;

                        const result = await Promise.resolve(
                            funcRef(eventContext)
                        );
                        preventDefault = result?.preventDefault;
                    }

                    if (preventDefault) {
                        actionDispatch({
                            type: "setWorkIndicator",
                            payload: false,
                        });
                    } else {
                        actionDispatch({
                            type: "setWorkIndicator",
                            payload: true,
                        });
                        const itemId = item.id;
                        const targetOption = dropResult.option as Option;
                        const update: any = {
                            [props.separatorMetadata.LogicalName]:
                                targetOption.Value,
                        };

                        if (
                            props.separatorMetadata.LogicalName === "statuscode"
                        ) {
                            update["statecode"] = targetOption.State;
                        }

                        await WebApiClient.Update({
                            entityName: props.metadata.LogicalName,
                            entityId: itemId,
                            entity: update,
                        });

                        actionDispatch({
                            type: "setWorkIndicator",
                            payload: false,
                        });
                        await props.refresh();
                    }
                } catch (ex) {
                    actionDispatch({
                        type: "setWorkIndicator",
                        payload: false,
                    });
                    Xrm.Navigation.openAlertDialog({
                        text: (ex as any).message,
                        title: "An error occured",
                    });
                }
            };

            asyncEnd(item, monitor);
        },
        collect: (monitor) => ({
            isDragging: monitor.isDragging(),
        }),
        canDrag: () => !props.config.preventTransitions,
    });

    const openInline = () => {
        props.openRecord({
            entityType: props.metadata.LogicalName,
            id: props.data[props.metadata?.PrimaryIdAttribute],
        });
    };

    console.log("state Name: " + appState.configViewName);

    console.log(
        `${props.metadata.LogicalName} tile ${
            props.data[props.metadata.PrimaryIdAttribute]
        } is rerendering`
    );

    React.useEffect(() => {
        if (!props.config.styleCallback) {
            return;
        }

        const executeStyleCallback = async () => {
            const styleCallbackResult = await Promise.resolve(
                accessFunc(props.config.styleCallback)({
                    data: props.data,
                    WebApiClient: WebApiClient,
                })
            );
            setOverriddenStyle(styleCallbackResult);
        };

        executeStyleCallback();
    }, [props.data, props.laneOption]);

    const selectRecord = (ev?: React.MouseEvent<HTMLElement, MouseEvent>) => {
        ev.stopPropagation();

        if (props.isSecondary) {
            return;
        }

        actionDispatch({
            type: "setSelectedRecords",
            payload: {
                [props.data[props.metadata.PrimaryIdAttribute]]:
                    !props.isSelected,
            },
        });
    };

    const useStyles = makeStyles({
        main: {
            display: "flex",
            flexDirection: "column",
            flexWrap: "wrap",
            columnGap: "16px",
            rowGap: "36px",
        },

        title: {
            ...shorthands.margin(0, 0, "12px"),
        },

        card: {
            width: "100%",
            maxWidth: "100%",
            height: "fit-content",
        },

        flex: {
            ...shorthands.gap("4px"),
            display: "flex",
            flexDirection: "row",
            alignItems: "center",
        },
        flexCards: {
            ...shorthands.gap("4px"),
            display: "flex",
            flexDirection: "row",
            alignItems: "start",
            flexWrap: "wrap",
        },

        appIcon: {
            ...shorthands.borderRadius("4px"),
            height: "32px",
        },

        caption: {
            color: tokens.colorNeutralForeground3,
        },

        cardFooter: {
            alignItems: "center",
            justifyContent: "space-between",
        },

        cardOnTrack: {
            borderLeftWidth: "3px",
            borderLeftStyle: "solid",
            borderLeftColor: "var(--colorPaletteGreenBackground3)",
        },
        cardDueToday: {
            borderLeftWidth: "3px",
            borderLeftStyle: "solid",
            borderLeftColor: "var(--colorPalettePeachBorderActive)",
        },
        cardOverdue: {
            borderLeftWidth: "3px",
            borderLeftStyle: "solid",
            borderLeftColor: "var(--colorPaletteRedBackground3)",
        },
        cardTest: {
            borderLeftWidth: "3px",
            borderLeftStyle: "solid",
            borderLeftColor: "#fec601",
        },
        cardReview: {
            borderLeftWidth: "3px",
            borderLeftStyle: "solid",
            borderLeftColor: "#008080",
        },
        cardClosed: {
            borderLeftWidth: "3px",
            borderLeftStyle: "solid",
            borderLeftColor: "var(--colorBrandBackground)",
        },
    });

    const styles = useStyles();

    const fetchTaskAssignees = async () => {
        const resourceAssignmentsUserExpanded = await new Promise<
            Array<{ [key: string]: any }>
        >((resolve) => {
            Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_resourceassignment",
                `?$filter=_msdyn_taskid_value eq '${props.data.msdyn_projecttaskid}'&$expand=msdyn_bookableresourceid($expand=UserId)`
            ).then(function onSuccess(result) {
                resolve(result.entities);
            });
        });

        let assignees: Array<{ [key: string]: any }> = [];
        if (resourceAssignmentsUserExpanded.length > 0) {
            resourceAssignmentsUserExpanded.forEach((item) => {
                const assigneeId =
                    item["msdyn_bookableresourceid"]["UserId"]["systemuserid"];
                const assigneeEmail =
                    item["msdyn_bookableresourceid"]["UserId"][
                        "internalemailaddress"
                    ];
                const assigneeFullname =
                    item["msdyn_bookableresourceid"]["UserId"]["fullname"];
                assignees.push({
                    id: assigneeId,
                    name: assigneeFullname,
                    email: assigneeEmail,
                });
            });

            return assignees;
        }
        return [];
    };
    React.useEffect(() => {
        fetchTaskAssignees().then((assignees) => {
            setAssignees(assignees);
        });
    }, []);

    const fetchProject = async () => {
        const project = await new Promise<{ [key: string]: any }>((resolve) => {
            Xrm.WebApi.retrieveRecord(
                "msdyn_project",
                `${props.data._msdyn_project_value}`,
                "?$select=ddsol_projectimage_url,_msdyn_projectmanager_value,_proj_manager_value"
            ).then(function onSuccess(result) {
                resolve(result);
            });
        });
        const imgRelativeUrl = project.ddsol_projectimage_url;
        const environmentUrl = Xrm.Utility.getGlobalContext().getClientUrl();
        const imgUrl = environmentUrl + imgRelativeUrl;
        const managers = [
            project._msdyn_projectmanager_value,
            project._proj_manager_value,
        ];
        return { imgUrl: imgUrl, managers: managers };
    };
    React.useEffect(() => {
        fetchProject().then((obj) => {
            setProjectImageUrl(obj.imgUrl);
            setManagers(obj.managers);
        });
    }, []);

    const fetchTaskAttachments = async () => {
        const taskAttachments = await new Promise<
            Array<{ [key: string]: any }>
        >((resolve) => {
            Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_projecttaskattachment",
                `?$filter=_msdyn_task_value eq '${props.data.msdyn_projecttaskid}'&$select=msdyn_linkuri,msdyn_name`
            ).then(function onSuccess(result) {
                resolve(result.entities);
            });
        });

        let attachments: Array<{ [key: string]: any }> = [];
        if (taskAttachments.length > 0) {
            taskAttachments.forEach((item) => {
                const attachmentName = item["msdyn_name"];
                const attachmentUrl = item["msdyn_linkuri"];
                attachments.push({
                    attachmentName: attachmentName,
                    attachmentUrl: attachmentUrl,
                });
            });

            return attachments;
        }
        return [];
    };
    React.useEffect(() => {
        fetchTaskAttachments().then((attachments) => {
            setAttachments(attachments);
        });
    }, []);

    const renderSwitch = () => {
        switch (appState.configViewName) {
            case "CsCostingSheetCostingEngineerConfig":
                return (
                    <CostingSheetContent
                        styles={styles}
                        data={props.data}
                        openInline={openInline}
                    />
                );
            case "CsPlantPackageCeConfig":
                return (
                    <PlantPackageCE
                        styles={styles}
                        data={props.data}
                        completions={completions}
                        openInline={openInline}
                    />
                );
            case "CsPlantPackageKamConfig":
                return (
                    <PlantPackageKam
                        styles={styles}
                        data={props.data}
                        completions={completions}
                        openInline={openInline}
                    />
                );
            case "CsPlantPackageLogConfig":
                return (
                    <PlantPackageLog
                        styles={styles}
                        data={props.data}
                        completions={completions}
                        openInline={openInline}
                    />
                );
            case "CsPlantPackagePoConfig":
                return (
                    <PlantPackagePo
                        styles={styles}
                        data={props.data}
                        completions={completions}
                        openInline={openInline}
                    />
                );
            default:
                return <></>;
        }
    };

    return (
        <div onClick={selectRecord} ref={drag}>
            <FluentProvider theme={webLightTheme}>
                <CardFluent
                    className={mergeClasses(styles.card, styles.cardClosed)}
                >
                    {renderSwitch}
                </CardFluent>
            </FluentProvider>
        </div>
    );
};

const isDataEqual = (a: any, b: any) => {
    if (Object.keys(a).length != Object.keys(b).length) {
        return false;
    }

    if (
        Object.keys(a).some((k) => {
            const value = a[k];
            return b[k] !== value;
        })
    ) {
        return false;
    }

    return true;
};

export const Tile = React.memo(TileRender, (a, b) => {
    if (a.borderColor != b.borderColor) {
        return false;
    }

    if (a.cardForm != b.cardForm) {
        return false;
    }

    if (a.dndType != b.dndType) {
        return false;
    }

    if (a.laneOption != b.laneOption) {
        return false;
    }

    if (a.metadata != b.metadata) {
        return false;
    }

    if (a.searchText != b.searchText) {
        return false;
    }

    if (a.style != b.style) {
        return false;
    }

    if ((a.notifications || []).length != (b.notifications || []).length) {
        return false;
    }

    if ((a.subscriptions || []).length != (b.subscriptions || []).length) {
        return false;
    }

    if (a.isSelected != b.isSelected) {
        return false;
    }

    const secondaryNotificationsA = Object.keys(
        a.secondaryNotifications || {}
    ).reduce((all, cur) => [...all, ...a.secondaryNotifications[cur]], []);
    const secondaryNotificationsB = Object.keys(
        b.secondaryNotifications || {}
    ).reduce((all, cur) => [...all, ...b.secondaryNotifications[cur]], []);

    if (secondaryNotificationsA.length != secondaryNotificationsB.length) {
        return false;
    }

    const secondarySubscriptionsA = Object.keys(
        a.secondarySubscriptions || {}
    ).reduce((all, cur) => [...all, ...a.secondarySubscriptions[cur]], []);
    const secondarySubscriptionsB = Object.keys(
        b.secondarySubscriptions || {}
    ).reduce((all, cur) => [...all, ...b.secondarySubscriptions[cur]], []);

    if (secondarySubscriptionsA.length != secondarySubscriptionsB.length) {
        return false;
    }

    const secondaryDataA = a.secondaryData || [];
    const secondaryDataB = b.secondaryData || [];

    if (
        secondaryDataA.length != secondaryDataB.length ||
        secondaryDataA.some(
            (a, i) =>
                a.data.length != secondaryDataB[i].data.length ||
                a.data.some(
                    (d, j) => !isDataEqual(d, secondaryDataB[i].data[j])
                )
        )
    ) {
        return false;
    }

    return isDataEqual(a.data, b.data);
});

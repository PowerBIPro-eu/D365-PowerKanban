import * as React from "react";
import { useAppDispatch } from "../domain/AppState";
import { FieldRow } from "./FieldRow";
import { Metadata, Option, Attribute } from "../domain/Metadata";
import { CardForm } from "../domain/CardForm";
import { BoardLane } from "../domain/BoardLane";
import { Lane } from "./Lane";
import { ItemTypes } from "../domain/ItemTypes";
import { fetchSubscriptions, fetchNotifications, extractTextFromAttribute } from "../domain/fetchData";
import * as WebApiClient from "xrm-webapi-client";
import { useDrag, DragSourceMonitor } from "react-dnd";
import { FlyOutForm } from "../domain/FlyOutForm";
import { Notification } from "../domain/Notification";
import { BoardEntity } from "../domain/BoardViewConfig";
import { Subscription } from "../domain/Subscription";
import { useConfigState } from "../domain/ConfigState";
import { DisplayType, useActionDispatch } from "../domain/ActionState";
import { IButtonStyles } from "@fluentui/react/lib/Button";
import { ICardStyles } from '@uifabric/react-cards';
import { FetchUserAvatar } from "../domain/FetchUserInfo";
import { IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
import { Card as CardFluent, CardHeader } from "@fluentui/react-card";
import { Button } from "@fluentui/react-button";
import { FluentProvider } from "@fluentui/react-provider";
import { Menu, MenuItem, MenuList, MenuPopover, MenuTrigger } from "@fluentui/react-menu";
import { ProgressBar } from "@fluentui/react-progress";
import { Persona as PersonaFluent } from "@fluentui/react-persona";
import { Badge } from "@fluentui/react-badge";
import { Divider } from "@fluentui/react-divider";
import { Image as ImageFluent } from "@fluentui/react-image";
import { Text, Caption1, Subtitle2 } from "@fluentui/react-text";
import { Field } from "@fluentui/react-field";
// import { mergeClasses } from "@griffel/core/src/mergeClasses";
// import { makeStyles } from "@griffel/core/src/makeStyles";
import { makeStyles, mergeClasses } from "@fluentui/react-components";
import { shorthands } from "@griffel/core/";
import { tokens, webLightTheme } from "@fluentui/tokens";
import { Star16Filled, Alert16Filled, ArrowDown16Filled, ArrowRight24Regular, GanttChart24Regular, Important16Filled, CalendarAssistant20Regular, CalendarLtr20Regular, ChatMultiple16Regular, Circle16Filled, MeetNow24Regular, MoreHorizontal20Filled, ClipboardTask24Regular, ClockToolbox24Regular, Open16Regular } from "@fluentui/react-icons/lib/sizedIcons/chunk-ddsol";
import { title } from "process";
import { Link } from "@fluentui/react/lib/components/Link/Link";

interface TileProps {
    borderColor: string;
    cardForm: CardForm;
    config: BoardEntity;
    data: any;
    dndType?: string;
    laneOption?: Option;
    metadata: Metadata;
    notifications: Array<Notification>;
    searchText: string;
    secondaryData?: Array<BoardLane>;
    secondaryNotifications?: {[key: string]: Array<Notification>};
    secondarySubscriptions?: {[key: string]: Array<Subscription>};
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
    const appDispatch = useAppDispatch();
    const configState = useConfigState();
    const actionDispatch = useActionDispatch();
    const [overriddenStyle, setOverriddenStyle] = React.useState({} as ICardStyles);
    const [ personaUrl, setPersonaUrl ] = React.useState<string>(undefined);
    const [ projectImgUrl, setProjectImageUrl ] = React.useState<string>(undefined);
    const [ assignees, setAssignees ] = React.useState<Array<{ [key: string]: any }>>([]);
    const [ attachments, setAttachments ] = React.useState<Array<{ [key: string]: any }>>([]);
    const [ managers, setManagers ] = React.useState<Array<string>>([]);

    const secondaryConfig = configState.config.secondaryEntity;
    const secondaryMetadata = configState.secondaryMetadata[secondaryConfig ? secondaryConfig.logicalName : ""];
    const secondarySeparator = configState.secondarySeparatorMetadata;
    const stub = React.useRef(undefined);

    if (props.config.persona) {
        React.useEffect(() => {
            const personaAttribute = props.metadata.Attributes.find(a => a.LogicalName?.toLowerCase() === props.config.persona);

            if (!personaAttribute || personaAttribute.AttributeType !== "Owner") {
                return;
            }

            const ownerType = props.data[`_${props.config.persona}_value@Microsoft.Dynamics.CRM.lookuplogicalname`];
            const ownerId = props.data[`_${props.config.persona}_value`];

            if (ownerType !== "systemuser" || !ownerId) {
                return;
            }

            FetchUserAvatar(ownerId).then(url => {
                setPersonaUrl(url);
            });
        }, [ props.data[props.config.persona] ]);
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
            return actionDispatch({ type: "setWorkIndicator", payload: working });
        },
        data: props.data,
        WebApiClient: WebApiClient
    };

    const accessFunc = (identifier: string): Function => {
        const path = identifier.split(".");
        return path.reduce((all, cur) => !all ? undefined : (all as any)[cur], window) as any;
    };

    const [{ isDragging }, drag] = useDrag<{ id: string; sourceLane: Option, type: string } | undefined, undefined, {isDragging: boolean}>({
        item: { id: props.data[props.metadata.PrimaryIdAttribute], sourceLane: props.laneOption, type: props.dndType ?? ItemTypes.Tile } as any,
        end: (item: { id: string; sourceLane: Option } | undefined, monitor: DragSourceMonitor) => {
            const asyncEnd = async (item: { id: string; sourceLane: Option } | undefined, monitor: DragSourceMonitor) => {
                const dropResult = monitor.getDropResult();

                if (!dropResult || dropResult?.option?.Value == null || dropResult.option.Value === item.sourceLane.Value) {
                    return;
                }

                try {
                    let preventDefault = false;

                    if (props.config.transitionCallback) {
                        const eventContext = {
                            ...context,
                            target: dropResult.option
                        };

                        const funcRef = accessFunc(props.config.transitionCallback) as any;

                        const result = await Promise.resolve(funcRef(eventContext));
                        preventDefault = result?.preventDefault;
                    }

                    if (preventDefault) {
                        actionDispatch({ type: "setWorkIndicator", payload: false });
                    }
                    else {
                        actionDispatch({ type: "setWorkIndicator", payload: true });
                        const itemId = item.id;
                        const targetOption = dropResult.option as Option;
                        const update: any = { [props.separatorMetadata.LogicalName]: targetOption.Value };

                        if (props.separatorMetadata.LogicalName === "statuscode") {
                            update["statecode"] = targetOption.State;
                        }

                        await WebApiClient.Update({ entityName: props.metadata.LogicalName, entityId: itemId, entity: update });
                        
                        actionDispatch({ type: "setWorkIndicator", payload: false });
                        await props.refresh();
                    }
                } catch (ex) {
                    actionDispatch({ type: "setWorkIndicator", payload: false });
                    Xrm.Navigation.openAlertDialog({ text: (ex as any).message, title: "An error occured" });
                }
            };

            asyncEnd(item, monitor);
        },
        collect: (monitor) => ({
          isDragging: monitor.isDragging()
        }),
        canDrag: () => !props.config.preventTransitions
    });

    const opacity = isDragging ? 0.4 : 1;

    const setSelectedRecord = () => {
        actionDispatch({ type: "setSelectedRecordDisplayType", payload: DisplayType.recordForm });
        actionDispatch({ type: "setSelectedRecord", payload: { entityType: props.metadata.LogicalName, id: props.data[props.metadata?.PrimaryIdAttribute] } });
    };

    const showNotifications = () => {
        actionDispatch({ type: "setSelectedRecordDisplayType", payload: DisplayType.notifications });
        actionDispatch({ type: "setSelectedRecord", payload: { entityType: props.metadata.LogicalName, id: props.data[props.metadata?.PrimaryIdAttribute] } });
    };

    const showNotificationsQuick = (ev?: React.MouseEvent<any>) => {
        if (ev) {
            ev.stopPropagation();
        }

        showNotifications();
    };

    const openInNewTab = () => {
        Xrm.Navigation.openForm({ entityName: props.metadata.LogicalName, entityId: props.data[props.metadata?.PrimaryIdAttribute], openInNewWindow: true });
    };

    const openInline = () => {
        props.openRecord({
            entityType: props.metadata.LogicalName,
            id: props.data[props.metadata?.PrimaryIdAttribute]
        });
    };

    const quickOpen = (ev?: React.MouseEvent<any>) => {
        if (ev) {
            ev.stopPropagation();
        }

        const handler = (props.config.defaultOpenHandler || "").toLowerCase();

        switch(handler) {
            case "modal":
                return openInModal();
            case "sidebyside":
                return setSelectedRecord();
            case "newwindow":
                return openInNewTab();
            default:
                return openInline();
        }
    };

    const openInModal = () => {
        const input : Xrm.Navigation.PageInputEntityRecord = {
			pageType: "entityrecord",
            entityName: props.metadata.LogicalName,
            entityId: props.data[props.metadata?.PrimaryIdAttribute]
        }

        const options : Xrm.Navigation.NavigationOptions = {
			target: 2,
			width: {
                value: 85,
                unit: "%"
            },
			position: 1
		};

        Xrm.Navigation.navigateTo(input, options)
        .then(() => props.refresh(), () => props.refresh());
    };

    const createNewSecondary = async () => {
        const parentLookup = configState.config.secondaryEntity.parentLookup;
        const data = {
            [parentLookup]: props.data[props.metadata.PrimaryIdAttribute],
            [`${parentLookup}type`]: props.metadata.LogicalName,
            [`${parentLookup}name`]: props.data[props.metadata.PrimaryNameAttribute]
        };

        const result = await Xrm.Navigation.openForm({ entityName: secondaryMetadata.LogicalName, useQuickCreateForm: true }, data);

        if (result && result.savedEntityReference) {
            props.refresh();
        }
    };

    const createSubscription = async (emailNotificationsEnabled: boolean, emailNotificationsSender: string = undefined) => {
        actionDispatch({ type: "setWorkIndicator", payload: true });

        await WebApiClient.Create({
            entityName: "oss_subscription",
            entity: {
                [`${props.config.subscriptionLookup}@odata.bind`]: `/${props.metadata.LogicalCollectionName}(${props.data[props.metadata.PrimaryIdAttribute].replace("{", "").replace("}", "")})`,
                "oss_emailnotificationsenabled": emailNotificationsEnabled,
                "oss_emailnotificationssender": emailNotificationsSender
            }
        });

        const subscriptions = await fetchSubscriptions(configState.config);
        appDispatch({ type: "setSubscriptions", payload: subscriptions });
        actionDispatch({ type: "setWorkIndicator", payload: false });
    };

    const subscribe = (): any => createSubscription(false);
    const subscribeWithEmail = (): any => createSubscription(true, props.config.emailNotificationsSender ? JSON.stringify(props.config.emailNotificationsSender) : undefined);

    const unsubscribe = () => {
        const asyncAction = async () => {
            actionDispatch({ type: "setWorkIndicator", payload: true });
            const subscriptionsToDelete = props.subscriptions.filter(s => s[`_${props.config.subscriptionLookup}_value`] === props.data[props.metadata.PrimaryIdAttribute]);

            await Promise.all(subscriptionsToDelete.map(s =>
                WebApiClient.Delete({
                    entityName: "oss_subscription",
                    entityId: s.oss_subscriptionid
                })
            ));

            const subscriptions = await fetchSubscriptions(configState.config);
            appDispatch({ type: "setSubscriptions", payload: subscriptions });
            actionDispatch({ type: "setWorkIndicator", payload: false });
        };
        
        asyncAction();
    };

    const clearNotifications = () => {
        const asyncAction = async () => {
            actionDispatch({ type: "setWorkIndicator", payload: true });
            const notificationsToDelete = props.notifications;

            await Promise.all(notificationsToDelete.map(s =>
                WebApiClient.Delete({
                    entityName: "oss_notification",
                    entityId: s.oss_notificationid
                })
            ));

            const notifications = await fetchNotifications(configState.config);
            appDispatch({ type: "setNotifications", payload: notifications });
            actionDispatch({ type: "setWorkIndicator", payload: false });
        };

        asyncAction();
    };

    const initCallBack = (identifier: string) => {
        return async () => {
            const funcRef = accessFunc(identifier) as any;
            return Promise.resolve(funcRef(context));
        };
    };

    const isSubscribed = props.subscriptions && props.subscriptions.length;
    const isMailSubscribed = isSubscribed && props.subscriptions.some(s => s.oss_emailnotificationsenabled);
    const hasNotifications = props.notifications && props.notifications.length;

    console.log(`${props.metadata.LogicalName} tile ${props.data[props.metadata.PrimaryIdAttribute]} is rerendering`);

    const customSplitButtonStyles: IButtonStyles = {
        splitButtonMenuButton: { backgroundColor: 'white', width: 28, border: 'none' },
        splitButtonMenuIcon: { fontSize: '7px' },
        splitButtonDivider: { backgroundColor: '#c8c8c8', width: 1, right: 26, position: 'absolute', top: 4, bottom: 4 },
        splitButtonContainer: {
          selectors: {
            ["@media screen and (-ms-high-contrast: active)"]: { border: 'none' },
          },
        },
      };

    const menuProps: IContextualMenuProps = {
        items: [
            {
                key: 'open',
                text: 'Open',
                iconProps: { iconName: 'Forward' },
                onClick: openInline
            },
            {
                key: 'openInSplitScreen',
                text: 'Open In Splitscreen',
                iconProps: { iconName: 'OpenPaneMirrored' },
                onClick: setSelectedRecord
            },
            {
                key: 'openInNewWindow',
                text: 'Open In New Window',
                iconProps: { iconName: 'OpenInNewWindow' },
                onClick: openInNewTab
            },
            {
                key: 'openInModal',
                text: 'Open In Modal',
                iconProps: { iconName: 'Picture' },
                onClick: openInModal
            },
            (secondaryConfig && secondaryMetadata
                ? {
                key: "createNewSecondary",
                text: `Create new ${secondaryMetadata.DisplayName.UserLocalizedLabel.Label}`,
                iconProps: { iconName: 'Add'},
                onClick: createNewSecondary
                }
                : null
            ),
            ...(props.config.customButtons && props.config.customButtons.length ? props.config.customButtons.map(c => ({key: c.id, text: c.label, iconProps: { iconName: c.icon.value }, onClick: initCallBack(c.callBack)})) : [])
        ],
    };

    const subscriptionMenuProps: IContextualMenuProps = {
        items: [
            {
                key: 'subscribe',
                text: 'Subscribe',
                iconProps: { iconName: 'Ringer' },
                onClick: subscribe
            },
            props.config.emailSubscriptionsEnabled
            ? {
                key: 'subscribeWithEmail',
                text: 'Subscribe with Email',
                iconProps: { iconName: 'Mail' },
                onClick: subscribeWithEmail
            }
            : undefined,
            {
                key: 'unsubscribe',
                text: 'Unsubscribe',
                iconProps: { iconName: 'RingerRemove' },
                onClick: unsubscribe
            },
            {
                key: 'markAsRead',
                text: 'Mark as read',
                iconProps: { iconName: 'Hide3' },
                onClick: clearNotifications
            },
            {
                key: 'showNotifications',
                text: 'Show Notifications',
                iconProps: { iconName: 'View' },
                onClick: showNotifications
            }
        ]
    };

    const iconName = isMailSubscribed
        ? (hasNotifications ? 'MailSolid' : 'Mail')
        : (isSubscribed ? (hasNotifications ? 'RingerSolid' : 'Ringer') : 'RingerOff');

    React.useEffect(() => {
        if (!props.config.styleCallback) {
            return;
        }

        const executeStyleCallback = async () => {
            const styleCallbackResult = await Promise.resolve(accessFunc(props.config.styleCallback)({ data: props.data, WebApiClient: WebApiClient }));
            setOverriddenStyle(styleCallbackResult);
        };

        executeStyleCallback();
    }, [props.data, props.laneOption]);

    const personaTitle = props.config.persona ? extractTextFromAttribute(props.data, props.config.persona) : props.data[props.metadata.PrimaryNameAttribute];
    const headerData = <div style={{display: "flex", flex: "1", overflow: "auto", flexDirection: "column", color: "#666666" }}>
          { props.cardForm.parsed.header.rows.map((r, i) => <div key={`headerRow_${props.data[props.metadata.PrimaryIdAttribute]}_${i}`} style={{ flex: "1" }}><FieldRow searchString={props.searchText} type="header" metadata={props.metadata} data={props.data} cells={r.cells} /></div>) }                  
    </div>;

    const selectRecord = (ev?: React.MouseEvent<HTMLElement, MouseEvent>) => {
        ev.stopPropagation();

        if(props.isSecondary) {
            return;
        }

        actionDispatch({ type: "setSelectedRecords", payload: { [props.data[props.metadata.PrimaryIdAttribute]]: !props.isSelected } });
    };

    /* ----------------- DDSol Customizations ------------------ */
    const useStyles = makeStyles({
        main: {
          display: "flex",
          flexDirection: "column",
          flexWrap: "wrap",
          columnGap: "16px",
          rowGap: "36px"
        },
      
        title: {
          ...shorthands.margin(0, 0, "12px")
        },
      
        card: {
          width: "100%",
          maxWidth: "100%",
          height: "fit-content"
        },
      
        flex: {
          ...shorthands.gap("4px"),
          display: "flex",
          flexDirection: "row",
          alignItems: "center"
        },
        flexCards: {
          ...shorthands.gap("4px"),
          display: "flex",
          flexDirection: "row",
          alignItems: "start",
          flexWrap: "wrap"
        },
      
        appIcon: {
          ...shorthands.borderRadius("4px"),
          height: "32px"
        },
      
        caption: {
          color: tokens.colorNeutralForeground3
        },
      
        cardFooter: {
          alignItems: "center",
          justifyContent: "space-between"
        },
      
        cardOnTrack: {
          borderLeftWidth: "3px",
          borderLeftStyle: "solid",
          borderLeftColor: "var(--colorPaletteGreenBackground3)"
        },
        cardDueToday: {
          borderLeftWidth: "3px",
          borderLeftStyle: "solid",
          borderLeftColor: "var(--colorPalettePeachBorderActive)"
        },
        cardOverdue: {
          borderLeftWidth: "3px",
          borderLeftStyle: "solid",
          borderLeftColor: "var(--colorPaletteRedBackground3)"
        },
        cardTest: {
          borderLeftWidth: "3px",
          borderLeftStyle: "solid",
          borderLeftColor: "#fec601"
        },
        cardReview: {
          borderLeftWidth: "3px",
          borderLeftStyle: "solid",
          borderLeftColor: "#008080"
        },
        cardClosed: {
          borderLeftWidth: "3px",
          borderLeftStyle: "solid",
          borderLeftColor: "var(--colorBrandBackground)"
        }
    });

    const styles = useStyles();

    const getPotentialRewardStars = (rating: number) => {
        let stars = [];
        for (let index = 0; index < rating; index++) {
            stars.push(<Star16Filled />)
        }
        return stars;
    }

    const getPriorityBadge = (priority: number) => {
        switch (priority) {
            case 1:
                return <Badge
                    color="danger"
                    shape="rounded"
                    size="large"
                    icon={<Alert16Filled />}
                    title="Priority"
                >
                    Urgent
                </Badge>
            case 3:
                return <Badge
                    color="severe"
                    shape="rounded"
                    size="large"
                    icon={<Important16Filled />}
                    title="Priority"
                >
                    Important
                </Badge>
            case 5:
                return <Badge
                    color="success"
                    shape="rounded"
                    size="large"
                    icon={<Circle16Filled />}
                    title="Priority"
                >
                    Medium
                </Badge>
            case 9:
                return <Badge
                    color="informative"
                    shape="rounded"
                    size="large"
                    icon={<ArrowDown16Filled />}
                    title="Priority"
                >
                    Low
                </Badge>
            default:
                return <Badge
                color="informative"
                shape="rounded"
                size="large"
                title="Priority"
            >
                Unknown Priority!
            </Badge>
        }
    }

    const getProgressBarMessage = (effortSpent: number, effort: number) => {
        return `Effort: ${effortSpent} out of expected ${effort} hours used`
    }

    const openTaskConversation = (url: string) => {
        window.open(url);
    }

    function startTaskConversation() {
        Xrm.Utility.showProgressIndicator("Starting Task Conversation in Teams");
        Xrm.WebApi.retrieveMultipleRecords(
            "environmentvariabledefinition",
            `?$filter=schemaname eq 'ddsol_PATriggerURLStartTaskConversation'&$select=environmentvariabledefinitionid&$expand=environmentvariabledefinition_environmentvariablevalue($select=value)`
        ).then(function onSucces(result) {
            const Url_new =
                result.entities[0]
                    .environmentvariabledefinition_environmentvariablevalue[0]
                    .value;
            const Data = JSON.stringify({
                taskId: props.data.msdyn_projecttaskid,
            });
            let req = new XMLHttpRequest();
            req.open("POST", Url_new, true);
            req.setRequestHeader("Content-Type", "application/json");
            req.send(Data);
            req.onload = function (data) {
                Xrm.Utility.closeProgressIndicator();
                if (req.status !== 200) {
                    var confirmOptions = { height: 200, width: 450 };
                    Xrm.Navigation.openErrorDialog({message:"Error while starting task conversation."});
                } else {
                    console.log(req.response);
                    const response = JSON.parse(req.response);
                    const messageURL = response.messageLink;
                    Xrm.Navigation.openConfirmDialog(
                        {
                            title: "Task Conversation Started!",
                            text: "Task conversation has been created. Open the conversation and post your questions as a reply please.",
                            cancelButtonLabel: "Stay in app",
                            confirmButtonLabel: "Open conversation"
                        },
                        confirmOptions
                    ).then(function(data) {
                        // formContext.data.refresh();
                        // formContext.ui.refreshRibbon();
                        if (data.confirmed === true) {
                            Xrm.Navigation.openUrl(messageURL);
                        }
                    });
                }
            };
            req.onerror = function () {
                Xrm.Utility.closeProgressIndicator();
                console.log(req.response);
            };
        });
    }

    const getMessageMenuItem = () => {
        return props.data.ddsol_teamsmsglink ?
        <MenuItem icon={<ChatMultiple16Regular />} onClick={() => openTaskConversation(props.data.ddsol_teamsmsglink)}>
            Open Task Conversation
        </MenuItem> :
        <MenuItem icon={<ChatMultiple16Regular />} onClick={startTaskConversation}>
            Start Task Conversation
        </MenuItem>        
    }

    const messageMenuItem = getMessageMenuItem();

    async function OpenPreFilledActivityTrackerForm(isMeeting: boolean) {
        const currentDateTime = new Date();

        let parameters : {[key: string]: any} = {};

        parameters["ddsol_starttime"] = currentDateTime;
    
        // fill lookup fields
        parameters["ddsol_project"] = props.data._msdyn_project_value;
        parameters["ddsol_projectname"] = props.data["_msdyn_project_value@OData.Community.Display.V1.FormattedValue"];
        parameters["ddsol_projecttype"] = props.data["_msdyn_project_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

        parameters["ddsol_projecttask"] = props.data.msdyn_projecttaskid;
        parameters["ddsol_projecttaskname"] = props.data.msdyn_subject;
        parameters["ddsol_projecttasktype"] = "msdyn_projecttask";

        // if isMeeting tak checknut ci je otvorena aktivita a ukoncit ju, a otvorit form noveho activity trackera s activity category koordinacna porada
        if (isMeeting) {
            const userID = Xrm.Utility.getGlobalContext().userSettings.userId;
            const activeActivityTracker = await new Promise<Array<{ [key: string]: any }>>((resolve, reject) => {
                Xrm.WebApi.retrieveMultipleRecords(
                    "ddsol_activitytracker",
                    `?$filter=ddsol_endtime eq null and _ownerid_value eq '${userID}' &$select=ddsol_activitytrackerid,ddsol_activitytitle`
                ).then(function onSucces(result) {resolve(result.entities)})
            });

            // 1 aktivny activity tracker record for current user
            if (activeActivityTracker.length === 1) {
                const confirmResult = await new Promise((resolve, reject) => {
                    Xrm.Navigation.openConfirmDialog(
                        {title: "One ongoing Activity Tracker found",
                        text: `Ongoing Activity Tracker ${activeActivityTracker[0]["ddsol_activitytitle"]} has been found. End it now and start meeting activity tracker for task ${props.data.msdyn_subject}? Due to current Power Automate capabilities, meeting has to be started manually.`}
                        ).then(function onSuccess(result) {resolve(result.confirmed)})});
                if (confirmResult) {
                    Xrm.WebApi.updateRecord("ddsol_activitytracker", activeActivityTracker[0]["ddsol_activitytrackerid"], {"ddsol_endtime": currentDateTime});
                }
            }

            // viac aktivnych activity trackerov for current user, display alert
            if (activeActivityTracker.length > 1) {
                await Xrm.Navigation.openAlertDialog({title: "Multiple ongoing Activity Trackers!", text: "You have multiple ongoing Activity Trackers, close them manually. Activity Tracker meeting form will be opened."});
            }

            const meetingActivityCategoryName = "Koordinačná porada";
            const meetingActivityCategory = await new Promise<Array<{ [key: string]: any }>>((resolve, reject) => {
                Xrm.WebApi.retrieveMultipleRecords(
                "ddsol_activitycategory",
                `?$filter=ddsol_activitycategory eq '${meetingActivityCategoryName}'&$select=ddsol_activitycategoryid`
            ).then(function onSuccess(result) {resolve(result.entities)})});

            parameters["ddsol_activitycategory"] = meetingActivityCategory[0]["ddsol_activitycategoryid"];
            parameters["ddsol_activitycategoryname"] = meetingActivityCategoryName;
            parameters["ddsol_activitycategorytype"] = "ddsol_activitycategory";
        }
    
        // Define the table name to open the form  
        var entityFormOptions : {[key: string]: string}  = {};
        entityFormOptions["entityName"] = "ddsol_activitytracker";
    
        Xrm.Navigation.openForm(entityFormOptions, parameters)
    }

    const fetchTaskAssignees = async () => {
        const resourceAssignmentsUserExpanded = await new Promise<Array<{ [key: string]: any }>>((resolve, reject) => {
            Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_resourceassignment", `?$filter=_msdyn_taskid_value eq '${props.data.msdyn_projecttaskid}'&$expand=msdyn_bookableresourceid($expand=UserId)`
            ).then(function onSuccess(result) {resolve(result.entities)})
        });

        let assignees: Array<{ [key: string]: any }> = [];
        if (resourceAssignmentsUserExpanded.length > 0) {
            resourceAssignmentsUserExpanded.forEach(item => {
                const assigneeId = item["msdyn_bookableresourceid"]["UserId"]["systemuserid"];
                const assigneeEmail = item["msdyn_bookableresourceid"]["UserId"]["internalemailaddress"];
                const assigneeFullname = item["msdyn_bookableresourceid"]["UserId"]["fullname"];
                assignees.push({"id": assigneeId, "name": assigneeFullname, "email": assigneeEmail});
            });

            return assignees;
        }
        return [];
    }
    React.useEffect(() => { fetchTaskAssignees().then(assignees => {setAssignees(assignees)}) }, []);

    const fetchProject = async () => {
        const project = await new Promise<{ [key: string]: any }>((resolve, reject) => {
            Xrm.WebApi.retrieveRecord(
                "msdyn_project", `${props.data._msdyn_project_value}`, "?$select=ddsol_projectimage_url,_msdyn_projectmanager_value,_proj_manager_value"
            ).then(function onSuccess(result) {resolve(result)})
        });
        const imgRelativeUrl = project.ddsol_projectimage_url;
        const environmentUrl = Xrm.Utility.getGlobalContext().getClientUrl();
        const imgUrl = environmentUrl + imgRelativeUrl;
        const managers = [project._msdyn_projectmanager_value, project._proj_manager_value];
        return {imgUrl: imgUrl, managers: managers};
    }
    React.useEffect(() => {fetchProject().then(obj => {
        setProjectImageUrl(obj.imgUrl);
        setManagers(obj.managers);
    })}, []);

    const getAssigneeAvatars = React.useMemo(() => {
        if (assignees?.length > 1) {
            let assigneeAvatars: JSX.Element[] = [];
            assignees.forEach((item: { [key: string]: any }) => {
                assigneeAvatars.push(
                    <PersonaFluent
                        // name="TBD"
                        avatar={{
                            image: {
                            src:
                                `https://insys273.sharepoint.com/_layouts/15/userphoto.aspx?AccountName=${item.email}&Size=S`
                            },
                            title: item.name
                        }}
                    />
                );
            });
            return assigneeAvatars;
        }
        return <PersonaFluent
        name={assignees[0]?.name ?? ""}
        avatar={{
            image: {
            src:
                `https://insys273.sharepoint.com/_layouts/15/userphoto.aspx?AccountName=${assignees[0]?.email ?? ""}&Size=S`
            },
            title: assignees[0]?.name ?? ""
        }}
    />
    }, [assignees]);

    const fetchTaskAttachments = async () => {
        const taskAttachments = await new Promise<Array<{ [key: string]: any }>>((resolve, reject) => {
            Xrm.WebApi.retrieveMultipleRecords(
                "msdyn_projecttaskattachment", `?$filter=_msdyn_task_value eq '${props.data.msdyn_projecttaskid}'&$select=msdyn_linkuri,msdyn_name`
            ).then(function onSuccess(result) {resolve(result.entities)})
        });

        let attachments: Array<{ [key: string]: any }> = [];
        if (taskAttachments.length > 0) {
            taskAttachments.forEach(item => {
                const attachmentName = item["msdyn_name"];
                const attachmentUrl = item["msdyn_linkuri"];
                attachments.push({"attachmentName": attachmentName, "attachmentUrl": attachmentUrl});
            });

            return attachments;
        }
        return [];
    }
    React.useEffect(() => { fetchTaskAttachments().then(attachments => {setAttachments(attachments)}) }, []);

    const getAttachmentCards = React.useMemo(() => {
        const attachmentLogos = {
            "pdf": "https://spoppe-b.azureedge.net/files/fabric-cdn-prod_20211104.001/assets/item-types/32_1.5x/pdf.svg",
            "word": "https://spoppe-b.azureedge.net/files/fabric-cdn-prod_20211104.001/assets/item-types/32_1.5x/docx.svg",
            "excel": "https://spoppe-b.azureedge.net/files/fabric-cdn-prod_20211104.001/assets/item-types/32_1.5x/xlsx.svg",
            "onenote": "https://spoppe-b.azureedge.net/files/fabric-cdn-prod_20211104.001/assets/item-types/32_1.5x/onetoc.svg",
            "powerbi": "https://spoppe-b.azureedge.net/files/fabric-cdn-prod_20221015.001/assets/item-types/20/powerbi.svg",
            "file": "https://spoppe-b.azureedge.net/files/fabric-cdn-prod_20211104.001/assets/item-types/32_1.5x/txt.svg"
        }
        if (attachments.length > 0) {
            let attachmentsCards: JSX.Element[] = [];
            attachments.forEach((item: { [key: string]: any }) => {
                let attachmentImgUrl: string = "";
                switch (true) {
                    case item.attachmentUrl.includes('.pdf'):
                        attachmentImgUrl = attachmentLogos.pdf;
                    case item.attachmentUrl.includes('.docx'):
                        attachmentImgUrl = attachmentLogos.word;
                        break;
                    case item.attachmentUrl.includes('.xlsx'):
                        attachmentImgUrl = attachmentLogos.excel;
                        break;
                    case item.attachmentUrl.includes('OneNote.aspx'):
                        attachmentImgUrl = attachmentLogos.onenote;
                        break;
                    case item.attachmentUrl.includes('.pbix'):
                        attachmentImgUrl = attachmentLogos.powerbi;
                        break;
                    default:
                        attachmentImgUrl = attachmentLogos.file;
                        break;
                };
                attachmentsCards.push(
                    <Link href={item.attachmentUrl} target="_blank">
                        <CardFluent size="small" role="listitem">
                            <CardHeader
                            image={{
                                as: "img",
                                src: attachmentImgUrl,
                                alt: "Attachment type logo",
                                width: "30px",
                                height: "30px"
                            }}
                            header={<Text weight="semibold">{item.attachmentName}</Text>}
                            action={
                                <Button appearance="transparent" icon={<Open16Regular />} />
                            }
                            />
                        </CardFluent>
                    </Link>
                );
            });
            return attachmentsCards;
        }
    }, [attachments]);

    const checkIfDragable = React.useMemo(() => {
        let userId = Xrm.Utility.getGlobalContext().userSettings.userId.slice(1, -1).toLowerCase();
        const bucket = props.data["ddsol_kanbanbucket@OData.Community.Display.V1.FormattedValue"];
        if (bucket === "Review" || bucket === "Done") {
            return managers.filter(manager => {
                return manager === userId;
            }).length > 0;
        }
        return assignees.filter(assignee => {
            return assignee.id === userId;
        }).length > 0 || managers.filter(manager => {
            return manager === userId;
        }).length > 0;
    }, [assignees, managers]);
    
    return (
        <div onClick={selectRecord} ref={ checkIfDragable ? drag : stub}>
            <FluentProvider theme={webLightTheme}>
                <CardFluent className={mergeClasses(styles.card, styles.cardClosed)}>
                    <CardHeader
                    image={
                        <ImageFluent
                            alt="Project Logo"
                            src={projectImgUrl}
                            height={30}
                            width={30}
                            />
                    }
                    header={
                        <Subtitle2>
                        <b>{props.data.msdyn_subject}</b>
                        </Subtitle2>
                    }
                    action={
                        <Menu>
                            <MenuTrigger disableButtonEnhancement>
                                <Button
                                appearance="transparent"
                                icon={<MoreHorizontal20Filled />}
                                aria-label="More options"
                                />
                            </MenuTrigger>

                            <MenuPopover>
                                <MenuList>
                                    <MenuItem icon={<ClockToolbox24Regular />} onClick={() => {OpenPreFilledActivityTrackerForm(false)}}>
                                        New Activity Tracker
                                    </MenuItem>
                                    <MenuItem icon={<MeetNow24Regular />} onClick={() => {OpenPreFilledActivityTrackerForm(true)}}>Start Meeting Activity</MenuItem>
                                    <MenuItem icon={<ClipboardTask24Regular />} onClick={openInModal}>
                                        Open Task
                                    </MenuItem>
                                    {messageMenuItem}
                                </MenuList>
                            </MenuPopover>
                        </Menu>
                    }
                    />
                    <header
                    className={mergeClasses(styles.flex)}
                    style={{ flexWrap: "wrap" }}
                    >
                    <Badge
                        color="brand"
                        shape="rounded"
                        appearance="tint"
                        size="large"
                        icon={<GanttChart24Regular />}
                        title="Project"
                    >
                        {extractTextFromAttribute(props.data, "msdyn_project")}
                    </Badge>
                    <Badge
                        color="brand"
                        shape="rounded"
                        appearance="tint"
                        size="large"
                        title="Potential Reward Rating"
                    >
                        {getPotentialRewardStars(props.data.ddsol_potentialrewardrating)}
                    </Badge>
                    {getPriorityBadge(props.data.msdyn_priority)}
                    </header>

                    <div>
                    <Text block weight="semibold">
                        Task description
                    </Text>
                    <Caption1 block className={styles.caption}><div dangerouslySetInnerHTML={{__html: props.data.msdyn_description}}></div></Caption1>
                    </div>

                    <div
                    className={mergeClasses(styles.flex, styles.cardFooter)}
                    style={{ alignItems: "flex-start" }}
                    >
                    <div>
                        <Text block weight="semibold">
                        Assigned to
                        </Text>
                        {getAssigneeAvatars}
                    </div>
                    <div>
                        <Text block weight="semibold">
                        Finish
                        </Text>
                        <div className={styles.flex}>
                        <CalendarAssistant20Regular />
                        <Text>{new Date(props.data.msdyn_finish).toDateString()}</Text>
                        </div>
                    </div>
                    </div>

                    <Field
                    validationMessage={getProgressBarMessage(props.data.ddsol_totalworkduration ?? 0, props.data.msdyn_effort ?? 0)}
                    validationState="none"
                    >
                    <ProgressBar
                        color="brand"
                        shape="rounded"
                        thickness="large"
                        value={(props.data.ddsol_totalworkduration ?? 0 /props.data.msdyn_effort)}
                    />
                    </Field>

                    <div className={styles.flexCards}>
                        {getAttachmentCards}
                    </div>

                    <Divider />
                    <footer className={mergeClasses(styles.flex, styles.cardFooter)}>
                    <div className={styles.flex}>
                        <CalendarLtr20Regular />
                        <Caption1>
                        Created on <i>{new Date(props.data.createdon).toDateString()}</i>
                        </Caption1>
                    </div>
                    <div>
                        <Button
                        icon={<Open16Regular />}
                        size="small"
                        style={{ marginRight: "1rem" }}
                        onClick={openInModal}
                        >
                        Open
                        </Button>
                    </div>
                    </footer>
                </CardFluent>
            </FluentProvider>
        </div>
    );
};

const isDataEqual = (a: any, b: any) => {
    if (Object.keys(a).length != Object.keys(b).length) {
        return false;
    }

    if (Object.keys(a).some(k => {
        const value = a[k];
        return b[k] !== value;
    })) {
        return false;
    }

    return true;
}

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

    const secondaryNotificationsA = Object.keys(a.secondaryNotifications || {}).reduce((all, cur) => [...all, ...a.secondaryNotifications[cur]], []);
    const secondaryNotificationsB = Object.keys(b.secondaryNotifications || {}).reduce((all, cur) => [...all, ...b.secondaryNotifications[cur]], []);

    if (secondaryNotificationsA.length != secondaryNotificationsB.length) {
        return false;
    }

    const secondarySubscriptionsA = Object.keys(a.secondarySubscriptions || {}).reduce((all, cur) => [...all, ...a.secondarySubscriptions[cur]], []);
    const secondarySubscriptionsB = Object.keys(b.secondarySubscriptions || {}).reduce((all, cur) => [...all, ...b.secondarySubscriptions[cur]], []);

    if (secondarySubscriptionsA.length != secondarySubscriptionsB.length) {
        return false;
    }

    const secondaryDataA = a.secondaryData || [];
    const secondaryDataB = b.secondaryData || [];

    if (secondaryDataA.length != secondaryDataB.length || secondaryDataA.some((a, i) => a.data.length != secondaryDataB[i].data.length || a.data.some((d, j) => !isDataEqual(d, secondaryDataB[i].data[j])))) {
        return false;
    }

    return isDataEqual(a.data, b.data);
});
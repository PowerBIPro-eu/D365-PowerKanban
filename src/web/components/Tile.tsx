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
//import { Card, ICardStyles } from '@uifabric/react-cards';
import { ICardStyles } from '@uifabric/react-cards';
// import { IconButton } from "@fluentui/react/lib/Button";
// import { Persona, PersonaSize } from "@fluentui/react/lib/Persona";
import { FetchUserAvatar } from "../domain/FetchUserInfo";
import { IContextualMenuItem, IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
import { Badge, Caption1, Card as CardFluent, CardHeader, Subtitle2, Text as TextFluent, Persona as PersonaFluent, Divider, Button as ButtonFluent, makeStyles, shorthands, tokens, mergeClasses, FluentProvider, webLightTheme } from "@fluentui/react-components";
import { CalendarAssistant20Regular, CalendarLtr20Regular } from "@fluentui/react-icons/lib/sizedIcons/chunk-2";
import { Open16Regular } from "@fluentui/react-icons/lib/sizedIcons/chunk-8";

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
                value: 70,
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

	/*----------------------------------------------------------------------*/
	/*------------------ React Fluent UI v9 Test code ---------------------*/
	/*--------------------------------------------------------------------*/
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
		cardDue7Days: {
		  borderLeftWidth: "3px",
		  borderLeftStyle: "solid",
		  borderLeftColor: "var(--colorPaletteYellowBackground3)"
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
		cardClosed: {
		  borderLeftWidth: "3px",
		  borderLeftStyle: "solid",
		  borderLeftColor: "var(--colorBrandBackground)"
		},
	  
		bgOnTrack: {
		  backgroundColor: "var(--colorPaletteGreenBackground3)",
		  color: "white"
		},
		bgIn7Days: {
		  backgroundColor: "var(--colorPaletteYellowBackground3)",
		  color: "black"
		},
		bgToday: {
		  backgroundColor: "var(--colorPalettePeachBorderActive)",
		  color: "white"
		},
		bgOverdue: {
		  backgroundColor: "var(--colorPaletteRedBackground3)",
		  color: "white"
		},
		bgClosed: {
		  backgroundColor: "var(--colorBrandBackground)",
		  color: "white"
		}
	  });

	  const stylesFluent = useStyles();

	const getTaskTimeStatus = (dueDateStr: string, taskStatus: string) => {
		const today = new Date();
		const dueDate = new Date(dueDateStr);
		if (taskStatus === "Closed") {
			return "Closed"
		}
		if (today.getFullYear() === dueDate.getFullYear() &&
		today.getMonth() === dueDate.getMonth() &&
		today.getDate() === dueDate.getDate()) {
			return "Due Today";
		}
		const diffTime = Math.abs(dueDate.getTime() - today.getTime());
		if (Math.ceil(diffTime / (1000 * 60 * 60 * 24)) <= 7) {
			return "Due in 7 Days"
		}
		if (today > dueDate) {
			return "Overdue"
		}
		return "On Track"
	}

	const getCardClass = (taskStatus: string) => {
		switch (taskStatus) {
			case "Closed":
				return mergeClasses(stylesFluent.card, stylesFluent.cardClosed)
				break;
			case "Due Today":
				return mergeClasses(stylesFluent.card, stylesFluent.cardDueToday)
				break;
			case "Due in 7 Days":
				return mergeClasses(stylesFluent.card, stylesFluent.cardDue7Days)
				break;
			case "Overdue":
				return mergeClasses(stylesFluent.card, stylesFluent.cardOverdue)
				break;
			case "On Track":
				return mergeClasses(stylesFluent.card, stylesFluent.cardOnTrack)
				break;
			default:
				break;
		}
	}

	const getBadgeClass = (taskStatus: string) => {
		switch (taskStatus) {
			case "Closed":
				return mergeClasses(stylesFluent.bgClosed)
				break;
			case "Due Today":
				return mergeClasses(stylesFluent.bgToday)
				break;
			case "Due in 7 Days":
				return mergeClasses(stylesFluent.bgIn7Days)
				break;
			case "Overdue":
				return mergeClasses(stylesFluent.bgOverdue)
				break;
			case "On Track":
				return mergeClasses(stylesFluent.bgOnTrack)
				break;
			default:
				break;
		}
	}

	const getTaskBadges = () => {
		const stage = props.data["crf44_qmstage@OData.Community.Display.V1.FormattedValue"];
		const badgeFields = props.config.customConfigOptions[stage];
		const badges = badgeFields.map(getTaskBadge);
		return badges;
	}

	const getTaskBadge = (key: string) => {
		const valueRaw = JSON.parse(props.data.crf44_qmcategorization)[key];
		const value = valueRaw.charAt(0).toUpperCase() + valueRaw.slice(1);
		return (
			<Badge color="brand" shape="rounded" appearance="tint" size="small">
				{value}
			</Badge>
		)
	}

	const taskTimeStatus = getTaskTimeStatus(props.data.scheduledend, props.data["crf44_qmstatus@OData.Community.Display.V1.FormattedValue"]);
	const cardClass = getCardClass(taskTimeStatus);
	const badgeClass = getBadgeClass(taskTimeStatus);
	const badges = getTaskBadges();
    
    return (
        <div onClick={selectRecord} ref={ props.preventDrag ? stub : drag}>
            {/* <Card tokens={{ childrenGap: "5px" }} styles={{ root: { maxWidth: "auto", backgroundColor: "#fff", opacity, borderStyle: "solid", borderWidth: "1px", borderColor: "#d8d8d8", borderLeftColor: props.borderColor, borderLeftWidth: "3px", ...props.style, ...overriddenStyle, ...(props.isSelected ? { boxShadow: "0 0 1em deepskyblue" } : {}) }}}>
                <Card.Section styles={{root: { padding: "10px", borderBottom: "1px solid rgba(0,0,0,.125)" }}}>
                    <div style={{display: "flex", flexDirection: "column"}}>
                        <div style={{display: "flex", flexDirection: "row"}}>
                            { props.config.persona !== null 
                                ? <Persona title={personaTitle} imageUrl={personaUrl} imageAlt={personaTitle} styles={{root: { marginRight: "5px" } }} text={personaTitle} size={PersonaSize.size32}></Persona> 
                                : headerData
                            }
                            <div style={{ marginLeft: "auto" }}>
                                { props.config.notificationLookup && props.config.subscriptionLookup && 
                                    <>
                                        <IconButton
                                            id="notificationButton"
                                            styles={customSplitButtonStyles}
                                            iconProps={{ iconName: iconName, style: { color: hasNotifications ? "red" : "inherit" }}}
                                            split
                                            aria-roledescription="split button"
                                            menuProps={{ items: subscriptionMenuProps.items.filter(m => !!m) }}
                                            onClick={showNotificationsQuick}
                                        />
                                    </>
                                }
                                <IconButton
                                    id="moreButton"
                                    styles={customSplitButtonStyles}
                                    iconProps={{ iconName: 'Forward' }}
                                    split
                                    aria-roledescription="split button"
                                    menuProps={({ items: menuProps.items.filter(m => !!m) })}
                                    onClick={quickOpen}
                                />
                            </div>
                        </div>
                        { props.config.persona !== null && headerData }
                    </div>
                </Card.Section>
                <Card.Section styles={{ root: { padding: "10px" }}}>
                    <div style={{display: "flex", overflow: "auto", flexDirection: "column" }}>
                        { props.cardForm.parsed.body.rows.map((r, i) => <div key={`bodyRow_${props.data[props.metadata.PrimaryIdAttribute]}_${i}`} style={{ flex: "1" }}><FieldRow searchString={props.searchText} type="body" metadata={props.metadata} data={props.data} cells={r.cells} /></div>) }
                    </div>
                    { props.secondaryData &&
                    <div>
                        <div className="border-top my-3"></div>
                        <span style={{marginLeft: "5px", fontSize: "larger"}}>
                            {secondaryMetadata.DisplayCollectionName.UserLocalizedLabel.Label}
                        </span>
                        <IconButton iconProps={{iconName: "Add"}} style={{marginLeft: "5px"}} onClick={createNewSecondary}></IconButton>
                        <div id="flexContainer" style={{ display: "flex", flexDirection: "row", overflow: "auto" }}>
                            {
                                props.secondaryData.map(d => <Lane
                                refresh={props.refresh}
                                notifications={props.secondaryNotifications}
                                searchText={props.searchText}
                                subscriptions={props.secondarySubscriptions}
                                dndType={`${ItemTypes.Tile}_${props.data[props.metadata.PrimaryIdAttribute]}`}
                                key={`lane_${d.option?.Value ?? "fallback"}`}
                                minWidth="300px"
                                cardForm={props.selectedSecondaryForm}
                                metadata={secondaryMetadata}
                                lane={d}
                                config={secondaryConfig}
                                separatorMetadata={secondarySeparator}
                                openRecord={props.openRecord}
                                isSecondaryLane />)
                            }
                        </div>
                    </div>
                    }
                </Card.Section>
                <Card.Section styles={{ root: { backgroundColor: "#efefef", padding: "10px", borderTop: "1px solid rgba(0,0,0,.125)" }}}>
                    <div style={{display: "flex", overflow: "auto", flexDirection: "column" }}>
                        { props.cardForm.parsed.footer.rows.map((r, i) => <div key={`footerRow_${props.data[props.metadata.PrimaryIdAttribute]}_${i}`} style={{ flex: "1" }}><FieldRow searchString={props.searchText} type="footer" metadata={props.metadata} data={props.data} cells={r.cells} /></div>) }
                    </div>
                </Card.Section>
            </Card> */}
			<FluentProvider theme={webLightTheme}>
				<CardFluent className={cardClass} /*{...props}*/>
					<CardHeader
						image={
							<Badge size="extra-large" className={badgeClass}>
								{props.data["crf44_qmstage@OData.Community.Display.V1.FormattedValue"]}
							</Badge>
						}
						header={
							<Subtitle2>
								<b>{props.data.subject}</b>
							</Subtitle2>
						}
						action={
						<Badge shape="rounded" className={badgeClass}>
							{taskTimeStatus}
						</Badge>
						}
					/>
					<header
						className={mergeClasses(stylesFluent.flex)}
						style={{ flexWrap: "wrap" }}
					>
						{badges}
					</header>

					<div>
						<TextFluent block weight="semibold">
							Action description
						</TextFluent>
						<Caption1 block className={stylesFluent.caption}>
							{props.data.description ?? "No additional action description"}
						</Caption1>
					</div>

					<div
						className={mergeClasses(stylesFluent.flex, stylesFluent.cardFooter)}
						style={{ alignItems: "flex-start" }}
					>
						<div>
						<TextFluent block weight="semibold">
							Responsible
						</TextFluent>
						<PersonaFluent
							name={props.data["_ownerid_value@OData.Community.Display.V1.FormattedValue"]}
							avatar={{
							image: {
								src:
								"https://magna.sharepoint.com/_layouts/15/userphoto.aspx?AccountName=insys.juricek@magna.com&Size=S"
							}
							}}
							// {...props}
						/>
						</div>
						<div>
						<TextFluent block weight="semibold">
							Due Date
						</TextFluent>
						<div className={stylesFluent.flex}>
							<CalendarAssistant20Regular />
							<TextFluent>{props.data["scheduledend@OData.Community.Display.V1.FormattedValue"].split(' ')[0]}</TextFluent>
						</div>
						</div>
					</div>

					<Divider />
					<footer className={mergeClasses(stylesFluent.flex, stylesFluent.cardFooter)}>
						<div className={stylesFluent.flex}>
						<CalendarLtr20Regular />
						<Caption1>
							Created on <i>{props.data["createdon@OData.Community.Display.V1.FormattedValue"]}</i>
						</Caption1>
						</div>
						<ButtonFluent icon={<Open16Regular />} size="small" onClick={openInModal}>
							Open
						</ButtonFluent>
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
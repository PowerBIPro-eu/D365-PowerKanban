import * as React from "react";
import { useAppContext } from "../domain/AppState";

import { fetchData, refresh, fetchNotifications } from "../domain/fetchData";
import { NotificationTile } from "./NotificationTile";
import * as WebApiClient from "xrm-webapi-client";
import { FieldRow } from "./FieldRow";
import { useActionContext } from "../domain/ActionState";
import { useConfigState } from "../domain/ConfigState";

import { Card, ICardTokens, ICardSectionStyles, ICardSectionTokens } from '@uifabric/react-cards';
import { PrimaryButton, IconButton } from "@fluentui/react/lib/Button";
import { getSplitBorderButtonStyle, getSplitBorderContainerStyle } from "../domain/Internationalization";

interface NotificationListProps {
}

export const NotificationList = (props: NotificationListProps) => {
  const [ actionState, actionDispatch ] = useActionContext();
  const [ eventRecord, setEventRecord ] = React.useState(undefined);
  const configState = useConfigState();
  const [ appState, appDispatch ] = useAppContext();

  const notificationRecord = actionState.selectedRecord;
  const notifications = appState.notifications[actionState.selectedRecord.id] ?? [];
  const columns = Array.from(new Set(notifications.reduce((all, cur) => [...all, ...cur.parsed.updatedFields], [] as Array<string>)));
  const eventMeta = actionState.selectedRecord.entityType === configState.config.primaryEntity.logicalName ? configState.metadata : configState.secondaryMetadata[actionState.selectedRecord.entityType];

  React.useEffect(() => {
    const fetchEventRecord = async() => {
      const data = await WebApiClient.Retrieve({ entityName: actionState.selectedRecord.entityType, entityId: actionState.selectedRecord.id, queryParams: `?$select=${columns.join(",")}`, headers: [ { key: "Prefer", value: "odata.include-annotations=\"*\"" } ] });
      setEventRecord(data);
    };
    fetchEventRecord();
  }, []);

  const closeSideBySide = () => {
    actionDispatch({ type: "setSelectedRecord", payload: undefined });
  };

  const clearAndRefresh = async () => {
    actionDispatch({ type: "setWorkIndicator", payload: true });

    await Promise.all(notifications.map(s =>
        WebApiClient.Delete({
            entityName: "oss_notification",
            entityId: s.oss_notificationid
        })
    ));

    const newNotifications = await fetchNotifications(configState.config);
    appDispatch({ type: "setNotifications", payload: newNotifications });
    actionDispatch({ type: "setWorkIndicator", payload: false });
    closeSideBySide();
  };

  const openInNewTab = () => {
    Xrm.Navigation.openForm({ entityName: eventRecord.LogicalName, entityId: eventRecord.Id, openInNewWindow: true });
  };

  const borderStyle = getSplitBorderContainerStyle(appState);
  const borderButtonStyle = getSplitBorderButtonStyle(appState);

  return (
      <div style={{ ...borderStyle, position: "relative", width: "100%", height: "100%" }}>
        {/* <IconButton title="Close" iconProps={{iconName: "ChromeClose"}} onClick={closeSideBySide} style={{ ...borderButtonStyle, color: "white", backgroundColor: "#045999", position: "absolute", zIndex: 1, top: "calc(50% - 40px)", left: "-18px" }}></IconButton>
        <IconButton title="Mark as read and close" iconProps={{iconName: "Hide3"}}  onClick={clearAndRefresh} style={{ ...borderButtonStyle, color: "white", backgroundColor: "#045999", position: "absolute", zIndex: 1, top: "50%", left: "-18px" }}></IconButton>
        <IconButton title="Open in new window" iconProps={{iconName: "OpenInNewWindow"}}  onClick={openInNewTab} style={{ ...borderButtonStyle, color: "white", backgroundColor: "#045999", position: "absolute", zIndex: 1, top: "calc(50% + 40px)", left: "-18px" }}></IconButton> */}
        { eventRecord &&
          <div>
            <Card tokens={{childrenGap: "5px"}} styles={{ root: { maxWidth: "auto", minWidth: "auto", margin: "5px", backgroundColor: "#f8f9fa" }}}>
              <Card.Section styles={{ root: { padding: "5px" }}}>Current Data</Card.Section>
              <Card.Section styles={{ root: { padding: "5px", borderBottom: "1px solid rgba(0,0,0,.125)" }}}>
                <div style={{display: "flex", overflow: "auto", flexDirection: "column" }}>
                  {
                    columns.filter(c => ["createdby", "modifiedon", "modifiedby", "modifiedonbehalfby", eventMeta.PrimaryIdAttribute].every(s => s !== c)).map(c =>
                      <div key={`currentRecord_${c}`} style={{ minWidth: "200px", margin: "5px", flex: "1" }}>
                        <FieldRow type="footer" metadata={eventMeta} data={eventRecord} cells={[ { field: c } ]} />
                      </div>
                    )
                  }
                  </div>
              </Card.Section>
            </Card>
          </div>
        }
        <div style={{overflow: "auto"}}>
          <Card tokens={{childrenGap: "10px"}} styles={{ root: { maxWidth: "auto", minWidth: "auto", margin: "5px", padding: "10px", backgroundColor: "#f8f9fa" }}}>
            { notifications.map(n => <NotificationTile key={n.oss_notificationid} parent={notificationRecord} data={n}></NotificationTile>)}
          </Card>
        </div>
      </div>
  );
};
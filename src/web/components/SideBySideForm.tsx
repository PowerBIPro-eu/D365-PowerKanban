import * as React from "react";
import { useAppContext } from "../domain/AppState";

import { useActionContext } from "../domain/ActionState";
import { useConfigState } from "../domain/ConfigState";
import { getSplitBorderContainerStyle } from "../domain/Internationalization";

export const SideBySideForm = () => {
    const [appState, appDispatch] = useAppContext();
    const [actionState, actionDispatch] = useActionContext();
    const configState = useConfigState();

    const borderStyle = getSplitBorderContainerStyle(appState);

    return (
        <div
            style={{
                ...borderStyle,
                position: "relative",
                width: "100%",
                height: "100%",
            }}
        >
            <iframe
                style={{ width: "100%", height: "100%", border: 0 }}
                src={`/main.aspx?pagetype=entityrecord${
                    configState.appId ? "&appid=" + configState.appId : ""
                }&navbar=off&etn=${actionState.selectedRecord.entityType}&id=${
                    actionState.selectedRecord.id
                }`}
            ></iframe>
        </div>
    );
};

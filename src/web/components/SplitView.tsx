import { Modal } from "@fluentui/react/lib/Modal";
import { ProgressIndicator } from "@fluentui/react/lib/ProgressIndicator";
import * as React from "react";
import { DisplayType, useActionState } from "../domain/ActionState";
import { useAppDispatch } from "../domain/AppState";
import { useConfigState } from "../domain/ConfigState";
import { Board } from "./Board";
import { ConfigSelector } from "./ConfigSelector";
import { ExternalForm } from "./ExternalForm";
import { SideBySideForm } from "./SideBySideForm";

interface SplitViewProps {
    primaryDataIds?: Array<string>;
}

export const SplitView = (props: SplitViewProps) => {
    const actionState = useActionState();
    const configState = useConfigState();
    const appDispatch = useAppDispatch();

    React.useEffect(() => {
        appDispatch({
            type: "setPrimaryDataIds",
            payload: props.primaryDataIds,
        });
    }, [props.primaryDataIds]);

    return (
        <>
            <Modal isOpen={!!actionState.progressText}>
                <div style={{ padding: "10px" }}>
                    <div
                        style={{
                            textAlign: "center",
                            width: "100%",
                            fontSize: "large",
                        }}
                    >
                        {actionState.progressText}
                    </div>
                    <br />
                    <ProgressIndicator />
                </div>
            </Modal>
            {actionState.flyOutForm && <ExternalForm />}
            <ConfigSelector
                show={
                    actionState.configSelectorDisplayState ||
                    !configState.configId
                }
            />
            <div
                style={{
                    display: "flex",
                    width: "100%",
                    height: "100%",
                    backgroundColor: "#efefef",
                }}
            >
                <div
                    style={
                        actionState.selectedRecord
                            ? {
                                  minWidth: "600px",
                                  resize: "horizontal",
                                  overflow: "auto",
                              }
                            : { width: "100%" }
                    }
                >
                    {configState.configId && <Board />}
                </div>
                {!!actionState.selectedRecord &&
                    actionState.selectedRecordDisplayType ===
                        DisplayType.recordForm && (
                        <div style={{ minWidth: "400px", flex: "1" }}>
                            <SideBySideForm />
                        </div>
                    )}
            </div>
        </>
    );
};

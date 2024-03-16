import * as React from "react";
import { IInputs } from "../PowerKanban/generated/ManifestTypes";
import { ActionStateProvider } from "../domain/ActionState";
import { AppStateProvider } from "../domain/AppState";
import { ConfigStateProvider } from "../domain/ConfigState";
import { SplitView } from "./SplitView";

export interface AppProps {
    configId?: string;
    primaryEntityLogicalName?: string;
    primaryEntityId?: string;
    appId?: string;
    primaryDataIds?: Array<string>;
    pcfContext: ComponentFramework.Context<IInputs>;
    aditionalData: any;
    configViewName: string;
}

export const App: React.FC<AppProps> = (props) => {
    return (
        <AppStateProvider
            primaryDataIds={props.primaryDataIds}
            primaryEntityId={props.primaryEntityId}
            pcfContext={props.pcfContext}
            aditionalData={props.aditionalData}
            configViewName={props.configViewName}
        >
            <ActionStateProvider>
                <ConfigStateProvider
                    appId={props.appId}
                    configId={props.configId}
                    primaryEntityLogicalName={props.primaryEntityLogicalName}
                >
                    <SplitView primaryDataIds={props.primaryDataIds} />
                </ConfigStateProvider>
            </ActionStateProvider>
        </AppStateProvider>
    );
};

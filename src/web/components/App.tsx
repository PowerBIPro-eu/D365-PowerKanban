import * as React from "react";
import { AppStateProvider } from "../domain/AppState";
import { SplitView } from "./SplitView";
import { ActionStateProvider } from "../domain/ActionState";
import { ConfigStateProvider } from "../domain/ConfigState";
import { ErrorBoundary } from "./ErrorBoundary";
import { IInputs } from "../PowerKanban/generated/ManifestTypes";

export interface AppProps
{
  configId?: string;
  primaryEntityLogicalName?: string;
  primaryEntityId?: string;
  appId?: string;
  primaryDataIds?: Array<string>;
  pcfContext: ComponentFramework.Context<IInputs>;
}

export const App: React.FC<AppProps> = (props) => {
  return (
    // <ErrorBoundary>
      <AppStateProvider primaryDataIds={props.primaryDataIds} primaryEntityId={props.primaryEntityId} pcfContext={props.pcfContext}>
        <ActionStateProvider>
          <ConfigStateProvider appId={props.appId} configId={props.configId} primaryEntityLogicalName={props.primaryEntityLogicalName}>
            {/* <ErrorBoundary> */}
              <SplitView primaryDataIds={props.primaryDataIds} />
            {/* </ErrorBoundary> */}
          </ConfigStateProvider>
        </ActionStateProvider>
      </AppStateProvider>
    // </ErrorBoundary>
  );
};
import * as React from "react";
import { IInputs } from "../PowerKanban/generated/ManifestTypes";
import { AppProps } from "../components/App";
import { BoardLane } from "./BoardLane";
import { Notification } from "./Notification";
import { Subscription } from "./Subscription";

type Action =
    | { type: "setBoardData"; payload: Array<BoardLane> }
    | { type: "setSecondaryData"; payload: Array<any> }
    | {
          type: "setSubscriptions";
          payload: { [key: string]: Array<Subscription> };
      }
    | {
          type: "setNotifications";
          payload: { [key: string]: Array<Notification> };
      }
    | { type: "setPrimaryDataIds"; payload: Array<string> };

export type AppStateDispatch = (action: Action) => void;

export type AppStateProps = {
    boardData?: Array<BoardLane>;
    secondaryData?: Array<BoardLane>;
    subscriptions?: { [key: string]: Array<Subscription> };
    notifications?: { [key: string]: Array<Notification> };
    primaryDataIds?: Array<string>;
    primaryEntityId?: string;
    pcfContext?: ComponentFramework.Context<IInputs>;
    aditionalData?: any;
    configViewName?: string;
};

function stateReducer(state: AppStateProps, action: Action): AppStateProps {
    switch (action.type) {
        case "setBoardData": {
            return { ...state, boardData: action.payload };
        }
        case "setSecondaryData": {
            return { ...state, secondaryData: action.payload };
        }
        case "setSubscriptions": {
            return { ...state, subscriptions: action.payload };
        }
        case "setNotifications": {
            return { ...state, notifications: action.payload };
        }
        case "setPrimaryDataIds": {
            return { ...state, primaryDataIds: action.payload };
        }
    }
}

const AppState = React.createContext<AppStateProps | undefined>(undefined);
const AppDispatch = React.createContext<AppStateDispatch | undefined>(
    undefined
);

export const AppStateProvider: React.FC<AppProps> = (props) => {
    const [state, dispatch] = React.useReducer(stateReducer, props ?? {});

    return (
        <AppState.Provider value={state}>
            <AppDispatch.Provider value={dispatch}>
                {props.children}
            </AppDispatch.Provider>
        </AppState.Provider>
    );
};

export function useAppState() {
    const context = React.useContext(AppState);

    if (!context) {
        throw new Error("useAppState must be used within a state provider!");
    }

    return context;
}

export function useAppDispatch() {
    const context = React.useContext(AppDispatch);

    if (!context) {
        throw new Error("useAppDispatch must be used within a state provider!");
    }

    return context;
}

export function useAppContext(): [AppStateProps, AppStateDispatch] {
    return [useAppState(), useAppDispatch()];
}

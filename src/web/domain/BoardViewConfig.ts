import { FlyOutForm } from "./FlyOutForm";
import { EntityReference } from "xrm-webapi-client";
import { DisplayState } from "../components/Board";

export interface CustomButton {
    id: string;
    icon: { type: string; value: string; };
    label: string;
    callBack: string;
}

export interface BoardEntity {
    logicalName: string;
    swimLaneSource: string;
    hiddenLanes: Array<number>;
    visibleLanes: Array<number>;
    emailSubscriptionsEnabled: boolean;
    emailNotificationsSender: { Id: string; LogicalName: string; };
    styleCallback: string;
    transitionCallback: string;
    notificationLookup: string;
    subscriptionLookup: string;
    preventTransitions: boolean;
    customButtons: Array<CustomButton>;
    fitLanesToScreenWidth: boolean;
    hideCountOnLane: boolean;
    defaultOpenHandler: "inline" | "sidebyside" | "modal" | "newwindow";
    persona: string;
	customConfigOptions: {[key: string]: string[]};
}

export interface SecondaryEntity extends BoardEntity {
    parentLookup: string;
    hiddenViews: Array<string>;
    visibleViews: Array<string>;
    defaultView: string;
}

export interface Context {
    showForm: (form: FlyOutForm) => Promise<any>;
}

export interface PrimaryEntity extends BoardEntity {

}

export interface BoardViewConfig {
    primaryEntity: PrimaryEntity;
    secondaryEntity: SecondaryEntity;
    customScriptUrl: string;
    customStyleUrl: string;
    cachingEnabled: boolean;
    defaultDisplayState: DisplayState;
}
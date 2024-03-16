import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { App, AppProps } from "../components/App";
import * as WebApiClient from "xrm-webapi-client";

import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class PowerKanban
    implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;

    private config: any = null;

    /**
     * Empty constructor.
     */
    constructor() {}

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public async init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ) {
        this._notifyOutputChanged = notifyOutputChanged;
        this._context = context;
        this._container = container;
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public async updateView(
        context: ComponentFramework.Context<IInputs>
    ): Promise<void> {
        if (this._context.parameters.primaryDataSet.loading) {
            return;
        }

        // Only render once all primary IDs are settled
        if (
            (this._context.parameters.primaryDataSet.paging as any).pageSize !==
            5000
        ) {
            this._context.parameters.primaryDataSet.paging.setPageSize(5000);
            this._context.parameters.primaryDataSet.refresh();

            return;
        }

        if (!this.config) {
            const configName = this._context.parameters.configName.raw;
            this.config = !configName
                ? null
                : await WebApiClient.Retrieve({
                      entityName: "oss_powerkanbanconfig",
                      alternateKey: [
                          { property: "oss_uniquename", value: configName },
                      ],
                      queryParams: "?$select=oss_powerkanbanconfigid",
                  });
        }

        let aditionalData: Object;
        switch (this._context.parameters.configName.raw) {
            case "CsCostingSheetCostingEngineerConfig":
                aditionalData = await this.loadMaterials(
                    aditionalData,
                    "ddsol_cs_hose",
                    "hoses"
                );
                aditionalData = await this.loadMaterials(
                    aditionalData,
                    "ddsol_cs_pipe",
                    "pipes"
                );
                aditionalData = await this.loadMaterials(
                    aditionalData,
                    "ddsol_cs_injectionmaterial",
                    "injectionMaterials"
                );
                aditionalData = await this.loadMaterials(
                    aditionalData,
                    "ddsol_cs_cstool",
                    "tools"
                );
                aditionalData = await this.loadMaterials(
                    aditionalData,
                    "ddsol_cs_csoperatingstep",
                    "operatingSteps"
                );

                break;
            case "CsPlantPackageCeConfig":
                aditionalData = await this.loadPlantPackageCounts(
                    aditionalData,
                    "ddsol_cs_bippackageplantprice",
                    "bipPrices",
                    "ddsol_packageplant/ddsol_cs_packageplantid,ddsol_sts_progress"
                );
                aditionalData = await this.loadPlantPackageCounts(
                    aditionalData,
                    "ddsol_cs_costingsheet",
                    "costingSheets",
                    "ddsol_packageplant/ddsol_cs_packageplantid,ddsol_sts_costingsheet"
                );
                break;
            case "CsPlantPackageKamConfig":
                aditionalData = await this.loadPlantPackageCounts(
                    aditionalData,
                    "ddsol_cs_bippackageplantprice",
                    "bipPrices",
                    "ddsol_packageplant/ddsol_cs_packageplantid,ddsol_sts_progress"
                );
                aditionalData = await this.loadPlantPackageCounts(
                    aditionalData,
                    "ddsol_cs_costingsheet",
                    "costingSheets",
                    "ddsol_packageplant/ddsol_cs_packageplantid,ddsol_sts_costingsheet"
                );
                break;
            case "CsPlantPackageLogConfig":
                aditionalData = await this.loadPlantPackageCounts(
                    aditionalData,
                    "ddsol_cs_bippackageplantprice",
                    "bipPrices",
                    "ddsol_packageplant/ddsol_cs_packageplantid,ddsol_sts_progress"
                );
                aditionalData = await this.loadPlantPackageCounts(
                    aditionalData,
                    "ddsol_cs_costingsheet",
                    "costingSheets",
                    "ddsol_packageplant/ddsol_cs_packageplantid,ddsol_sts_costingsheet"
                );
                break;
            case "CsPlantPackagePoConfig":
                aditionalData = await this.loadPlantPackageCounts(
                    aditionalData,
                    "ddsol_cs_bippackageplantprice",
                    "bipPrices",
                    "ddsol_packageplant/ddsol_cs_packageplantid,ddsol_sts_progress"
                );
                aditionalData = await this.loadPlantPackageCounts(
                    aditionalData,
                    "ddsol_cs_costingsheet",
                    "costingSheets",
                    "ddsol_packageplant/ddsol_cs_packageplantid,ddsol_sts_costingsheet"
                );
                break;
            default:
                break;
        }

        const props: AppProps = {
            appId: (this._context as any).page?.appId,
            primaryEntityLogicalName:
                this._context.parameters.primaryDataSet.getTargetEntityType(),
            configId: this.config ? this.config.oss_powerkanbanconfigid : null,
            primaryEntityId: (context.mode as any).contextInfo.entityId,
            primaryDataIds:
                this._context.parameters.primaryDataSet.sortedRecordIds,
            pcfContext: this._context,
            aditionalData: aditionalData,
            configViewName: this.config
                ? this._context.parameters.configName.raw
                : "",
        };

        ReactDOM.render(React.createElement(App, props), this._container);
    }

    private async loadMaterials(
        aditionalData: Object,
        table: string,
        attributeName: string
    ) {
        try {
            let costingSheetIds =
                this._context.parameters.primaryDataSet.sortedRecordIds.join(
                    " or ddsol_costingsheet/ddsol_cs_costingsheetid eq "
                );
            const materials = await WebApiClient.Retrieve({
                entityName: table,
                queryParams: `?$apply=filter(ddsol_costingsheet/ddsol_cs_costingsheetid eq ${costingSheetIds})/groupby((ddsol_costingsheet/ddsol_cs_costingsheetid), aggregate($count as count))`,
            });
            aditionalData = {
                ...aditionalData,
                [attributeName]: materials,
            };
        } catch (error) {
            console.error("Error retrieving data:", error);
        }
        return aditionalData;
    }

    private async loadPlantPackageCounts(
        aditionalData: Object,
        table: string,
        attributeName: string,
        groupByParams: string
    ) {
        try {
            let plantPackageIds =
                this._context.parameters.primaryDataSet.sortedRecordIds.join(
                    " or ddsol_packageplant/ddsol_cs_packageplantid eq "
                );
            const items = await WebApiClient.Retrieve({
                entityName: table,
                queryParams: `?$apply=filter(ddsol_packageplant/ddsol_cs_packageplantid eq ${plantPackageIds})/groupby((${groupByParams}), aggregate($count as count))`,
            });
            aditionalData = { ...aditionalData, [attributeName]: items };
        } catch (error) {
            console.error("Error retrieving data:", error);
        }
        return aditionalData;
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this._container);
    }
}

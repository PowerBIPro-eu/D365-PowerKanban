import * as React from "react";
import { useAppContext } from "../domain/AppState";

import { fetchData, refresh } from "../domain/fetchData";
import { UserInputModal } from "./UserInputModalProps";
import { useActionContext } from "../domain/ActionState";
import { useConfigContext } from "../domain/ConfigState";
import * as WebApiClient from "xrm-webapi-client";
import { formatGuid } from "../domain/GuidFormatter";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";

interface ConfigSelectorProps {
    show: boolean;
}

export const ConfigSelector = (props: ConfigSelectorProps) => {
    const [ actionState, actionDispatch ] = useActionContext();
    const [ configState, configDispatch ] = useConfigContext();
    const [ configId, setConfigId ] = React.useState(undefined);
    const [ configs, setConfigs ] = React.useState([]);
    const [ makeDefault, setMakeDefault ] = React.useState(false);

    const yesCallBack = async() => {
        const id = configId;

        if (makeDefault) {
            const userId = formatGuid(Xrm.Page.context.getUserId());
            await WebApiClient.Update({ entityName: "systemuser", entityId: userId, entity: { oss_defaultboardid: id } });
        }

        configDispatch({ type: "setConfigId", payload: id });
    };

    const hideDialog = () => {
        actionDispatch({ type: "setConfigSelectorDisplayState", payload: false });
    };

    React.useEffect(() => {
        const fetchConfigs = async() => {
            const { value: data }: { value: Array<any> } = await WebApiClient.Retrieve({entityName: "oss_powerkanbanconfig", queryParams: `?$select=oss_uniquename,oss_value,oss_powerkanbanconfigid&$filter=oss_entitylogicalname eq '${configState.primaryEntityLogicalName}'&$orderby=oss_uniquename` });
            setConfigs(data);
        };

        fetchConfigs();
    }, []);

    const onSelection = (e: any) => {
        setConfigId(e.target.value);
        setMakeDefault(false);
    };

    const onMakeDefault = (e: any) => {
        setMakeDefault(e.target.value);
    };

    const setConfig = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => {
        setConfigId(item.key);
    };

    return (
        <UserInputModal okButtonDisabled={!configId} noCallBack={() => {}} yesCallBack={yesCallBack} finally={hideDialog} title={"Choose Board"} show={props.show}>
            <Dropdown
              id="configSelector"
              label={configs.find(c => c.oss_powerkanbanconfigid === configId)?.oss_uniquename}
              // eslint-disable-next-line react/jsx-no-bind
              onChange={setConfig}
              placeholder="Select config"
              options={ configs.map(c => ({ key: c.oss_powerkanbanconfigid, text: c.oss_uniquename })) }
            />
        </UserInputModal>
        // <div>Not Yet Implemented</div>
    );
};
import * as React from "react";

import { IBasePickerProps, ITag } from "@fluentui/react/lib/Pickers";
import * as WebApiClient from "xrm-webapi-client";
import { useActionContext } from "../domain/ActionState";
import { FlyOutField, FlyOutLookupField } from "../domain/FlyOutForm";
import { extractTextFromAttribute } from "../domain/fetchData";

export interface IExtendedTag extends ITag {
    data: { [key: string]: any };
}

export interface IGenericEntityPickerProps
    extends IBasePickerProps<IExtendedTag> {}

export const ExternalForm = () => {
    const [actionState, actionDispatch] = useActionContext();
    const [formData, setFormData] = React.useState({} as any);
    const [pickData, setPickData] = React.useState(
        {} as { [key: string]: Array<IExtendedTag> }
    );

    const fields: Array<[string, FlyOutField]> = Object.keys(
        actionState.flyOutForm.fields
    ).map((fieldId) => [fieldId, actionState.flyOutForm.fields[fieldId]]);

    React.useEffect(() => {
        const lookups = fields.filter(
            ([, field]) => field.type.toLowerCase() === "lookup"
        );

        lookups.forEach(async ([fieldId, field]) => {
            const lookup = field as FlyOutLookupField;
            const entityNameGroups =
                /<\s*entity\s*name\s*=\s*["']([a-zA-Z_0-9]+)["']\s*>/gim.exec(
                    lookup.fetchXml
                );

            if (!entityNameGroups || !entityNameGroups.length) {
                return;
            }

            const entityName = entityNameGroups[1];

            const data = await WebApiClient.Retrieve({
                fetchXml: lookup.fetchXml,
                entityName: entityName,
                returnAllPages: true,
                headers: [
                    { key: "Prefer", value: 'odata.include-annotations="*"' },
                ],
            });
            setPickData({
                ...pickData,
                [fieldId]: data.value.map(
                    (d: any) =>
                        ({ key: d[`${entityName}id`], data: d } as IExtendedTag)
                ),
            });
        });
    }, [actionState.flyOutForm.fields]);

    const getTextFromItemByKey = (item: IExtendedTag, displayField: string) =>
        extractTextFromAttribute(item.data, displayField);

    return <div>Not Implemented Yet</div>;
};

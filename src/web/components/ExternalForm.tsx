import * as React from "react";
import { useAppContext } from "../domain/AppState";

import { extractTextFromAttribute, fetchData, refresh } from "../domain/fetchData";
import { UserInputModal } from "./UserInputModalProps";
import { useActionContext } from "../domain/ActionState";
import { FlyOutField, FlyOutLookupField } from "../domain/FlyOutForm";
import { TextField } from "@fluentui/react/lib/TextField";
import { Separator } from '@fluentui/react/lib/Separator';
import { TagPicker, ITag, IBasePicker, IInputProps, IBasePickerSuggestionsProps, ISuggestionItemProps, IBasePickerProps, IPickerItemProps, TagItem, BasePicker } from "@fluentui/react/lib/Pickers";
import { Label } from "@fluentui/react/lib/Label";
import { Text } from "@fluentui/react/lib/Text";
import * as WebApiClient from "xrm-webapi-client";
import { IconButton } from "@fluentui/react/lib/Button";
import { Toggle } from '@fluentui/react/lib/Toggle';
import { DatePicker, defaultDatePickerStrings } from "@fluentui/react/lib/DatePicker";

interface ExternalFormProps {
}

export interface IExtendedTag extends ITag {
    data: { [key: string]: any };
}

export interface IGenericEntityPickerProps extends IBasePickerProps<IExtendedTag> { }

class GenericEntityPickerProps extends BasePicker<IExtendedTag, IGenericEntityPickerProps> { }

export const ExternalForm = (props: ExternalFormProps) => {
    const [ actionState, actionDispatch ] = useActionContext();
    const [ formData, setFormData ] = React.useState({} as any);
    const [ pickData, setPickData ] = React.useState({} as { [key: string]: Array<IExtendedTag> });

    const fields: Array<[string, FlyOutField]> = Object.keys(actionState.flyOutForm.fields)
        .map(fieldId => [ fieldId, actionState.flyOutForm.fields[fieldId]]);

    const noCallBack = () => {
        actionState.flyOutForm.resolve({
            cancelled: true
        });
    };

    const yesCallBack = () => {
        actionState.flyOutForm.resolve({
            cancelled: false,
            values: formData
        });
    };

    const hideDialog = () => {
        actionDispatch({ type: "setFlyOutForm", payload: undefined });
    };

    const onFieldChange = (e: any) => {
        const value = e.target.value;
        const id = e.target.id;

        setFormData({...formData, [id]: value });
    };

    const onFieldChangeToggle = (e: any, checked: any) => {
        const value = checked;
        const id = e.target.id;

        setFormData({...formData, [id]: value });
    };

    const onFieldChangeDate = (date: any, fieldId: any) => {
        const value = date;
        const id = fieldId;

        setFormData({...formData, [id]: value });
    };

    React.useEffect(() => {
        const lookups = fields.filter(([fieldId, field]) => field.type.toLowerCase() === "lookup");

        lookups.forEach(async ([fieldId, field]) => {
            const lookup = field as FlyOutLookupField;
            const entityNameGroups = /<\s*entity\s*name\s*=\s*["']([a-zA-Z_0-9]+)["']\s*>/gmi.exec(lookup.fetchXml);

            if (!entityNameGroups || !entityNameGroups.length) {
                return;
            }            

            const entityName = entityNameGroups[1];

            const data = await WebApiClient.Retrieve({
                fetchXml: lookup.fetchXml,
                entityName: entityName,
                returnAllPages: true,
                headers: [ { key: "Prefer", value: "odata.include-annotations=\"*\"" } ]
            });
            //setPickData({...pickData, [fieldId]: data.value.map((d: any) => ({ key: d[`${entityName}id`], data: d } as IExtendedTag)) });
            setPickData(prevPickData => ({
                ...prevPickData,
                [fieldId]: data.value.map((d: any) => ({ key: d[`${entityName}id`], data: d } as IExtendedTag))
            }));
        });
    }, [ actionState.flyOutForm.fields ]);

    const textField = (fieldId: string, field: FlyOutField) => (
        <TextField
        key={fieldId} id={fieldId}
        description={field.subtext}
        required={field.required}
        multiline={field.rows && field.rows > 1}
        rows={field.rows ?? 1}
        type={field.type}
        label={field.label}
        placeholder={field.placeholder}
        onChange={onFieldChange}
        defaultValue={field.defaultValue}
        />
    );

    /* DDSol CUSTOM INPUT FIELDS */
    const toggleField = (fieldId: string, field: FlyOutField) => (
        field.defaultValue ?
        <Toggle key={fieldId} id={fieldId} label={field.label} defaultChecked onText="Yes" offText="No" onChange={onFieldChangeToggle}/> :
        <Toggle key={fieldId} id={fieldId} label={field.label} onText="Yes" offText="No" onChange={onFieldChangeToggle}/>
        
    )

    const datePicker = (fieldId: string, field: FlyOutField) => (
        <>
            <DatePicker
                placeholder="Select a date..."
                ariaLabel="Select a date"
                // DatePicker uses English strings by default. For localized apps, you must override this prop.
                strings={defaultDatePickerStrings}
                onSelectDate={(date: any) => onFieldChangeDate(date, fieldId)}
                key={fieldId} id={fieldId} label={field.label}
            />
            <Text styles={{root: { color: "#666666" } }} variant="small">{field.defaultValue != null ? `Responsible entered: ${new Date(field.defaultValue).toLocaleDateString()}` : ""}</Text>
        </>
    )

    const onItemSelected = (fieldId: string, item: IExtendedTag) => {
		console.log(item);
        setFormData({ ...formData, [fieldId]: item?.key, [`${fieldId}Name`]: item?.data?.fullname });
    };

    const getTextFromItemByKey = (item: IExtendedTag, displayField: string) => extractTextFromAttribute(item.data, displayField);

    const getTextFromItem = (item: IExtendedTag, field: FlyOutLookupField) => getTextFromItemByKey(item, field.displayField?.toLowerCase()) ?? "(No Data)";

    const filterSelectedTags = (fieldId: string, filterText: string, tagList: IExtendedTag[], field: FlyOutLookupField): IExtendedTag[] => {
        const data = pickData[fieldId];

        if (!data || !data.length) {
            return [];
        }

        if (!filterText) {
            return data;
        }        

        return data.filter(d => [field.displayField, ...(field.secondaryFields ?? [])].some(f => (getTextFromItemByKey(d, f) ?? "").toLowerCase().indexOf(filterText.toLowerCase()) !== -1));
    };

    const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
        suggestionsHeaderText: 'Suggested records',
        noResultsFoundText: 'No records found',
    };

    const inputProps: IInputProps = {
        'aria-label': 'Tag picker'
    };

    const suggestionItem: (props: IExtendedTag, itemProps: ISuggestionItemProps<IExtendedTag>, field: FlyOutLookupField) => JSX.Element = (props, itemProps, field) => {
        return (
            <div style={{padding: "5px", width: "100%", textAlign: "left"}}>
                <Text block>{getTextFromItem(props, field)}</Text>
                { (field.secondaryFields ?? []).map(f => f.toLowerCase()).map(f => [f, getTextFromItemByKey(props, f)]).filter(f => !!f[1]).map(f => <Text block key={f[0]} styles={{root: { color: "#666666" } }} variant="small">{f[1]}</Text>) }
            </div>
        );
    };

    const selectedItem: (props: IPickerItemProps<IExtendedTag>, field: FlyOutLookupField) => JSX.Element = (props, field) => {
        return (
            <TagItem index={0} onRemoveItem={props.onRemoveItem} item={props.item}>{getTextFromItem(props.item, field)}</TagItem>
        );
    };

    const lookupField = (fieldId: string, field: FlyOutLookupField) => (
        field.defaultValue ?
        <>
            <Label required={!!field.required}>{field.label}</Label>
            <GenericEntityPickerProps
                key={fieldId}
                removeButtonAriaLabel="Remove"
                onRenderItem={(props) => selectedItem(props, field)}
                onRenderSuggestionsItem={(props, itemProps) => suggestionItem(props, itemProps, field)}
                onResolveSuggestions={(filter: string, selectedItems?: IExtendedTag[]) => filterSelectedTags(fieldId, filter, selectedItems, field)}
                onChange={(items) => onItemSelected(fieldId, items && items.length ? items[0] : null)}
                onEmptyResolveSuggestions={(selectedItems?: IExtendedTag[]) => filterSelectedTags(fieldId, "", selectedItems, field)}
                onRemoveSuggestion={() => onItemSelected(fieldId, null)}
                getTextFromItem={(item: IExtendedTag) => getTextFromItem(item, field)}
                pickerSuggestionsProps={pickerSuggestionsProps}
                itemLimit={1}
                inputProps={inputProps}
                defaultSelectedItems={field.defaultSelectedId === null ? null :[{
                    key: field.defaultSelectedId,
                    name: field.defaultSelectedName,
                    data: { 
						fullname: field.defaultSelectedName,
						internalemailaddress: field.defaultSelectedName,
						ownerid: field.defaultSelectedId,
						systemuserid: field.defaultSelectedId
					}
                }]}

            />
            { field.subtext && <Text styles={{root: { color: "#666666" } }} variant="small">{field.subtext}</Text> }
        </> :
        <>
            <Label required={!!field.required}>{field.label}</Label>
            <GenericEntityPickerProps
                key={fieldId}
                removeButtonAriaLabel="Remove"
                onRenderItem={(props) => selectedItem(props, field)}
                onRenderSuggestionsItem={(props, itemProps) => suggestionItem(props, itemProps, field)}
                onResolveSuggestions={(filter: string, selectedItems?: IExtendedTag[]) => filterSelectedTags(fieldId, filter, selectedItems, field)}
                onChange={(items) => onItemSelected(fieldId, items && items.length ? items[0] : null)}
                onEmptyResolveSuggestions={(selectedItems?: IExtendedTag[]) => filterSelectedTags(fieldId, "", selectedItems, field)}
                onRemoveSuggestion={() => onItemSelected(fieldId, null)}
                getTextFromItem={(item: IExtendedTag) => getTextFromItem(item, field)}
                pickerSuggestionsProps={pickerSuggestionsProps}
                itemLimit={1}
                inputProps={inputProps}
            />
            { field.subtext && <Text styles={{root: { color: "#666666" } }} variant="small">{field.subtext}</Text> }
        </>
    );

    const triggerChangeOnLoad = (fields: {[key: string]: FlyOutField }) => {
        console.log('trigering change events when flyoutform loads');
        Object.keys(fields).forEach(fieldId => document.querySelector(`#${fieldId}`).dispatchEvent(new Event('change')));
    }

    return (
        <UserInputModal onLoad={() => triggerChangeOnLoad(actionState.flyOutForm.fields)} okButtonDisabled={!Object.keys(actionState.flyOutForm.fields).every(fieldId => !actionState.flyOutForm.fields[fieldId].required || !!formData[fieldId] || actionState.flyOutForm.fields[fieldId].defaultValue)} noCallBack={noCallBack} yesCallBack={yesCallBack} finally={hideDialog} title={actionState.flyOutForm?.title} show={!!actionState.flyOutForm}>
            <h4>Action</h4>
            <Text>{actionState.flyOutForm?.taskTitle}</Text>
            {
                actionState.flyOutForm?.taskDescription ?
                <>
                    <h4>Action Description</h4>
                    <Text>{actionState.flyOutForm?.taskDescription}</Text>
                </> : <></>
            }            
            <Separator>Action Responses</Separator>
            {Object.keys(actionState.flyOutForm.fields).map(fieldId => [ fieldId, actionState.flyOutForm.fields[fieldId]] as [string, FlyOutField]).map((
                [fieldId, field]) => 
                    field.type.toLowerCase() === "lookup" ? lookupField(fieldId, field as FlyOutLookupField)
                    : field.type.toLowerCase() === "toggle" ? toggleField(fieldId, field)
                    : field.type.toLowerCase() === "date" ? datePicker(fieldId, field) : textField(fieldId, field))}
        </UserInputModal>
    );
};
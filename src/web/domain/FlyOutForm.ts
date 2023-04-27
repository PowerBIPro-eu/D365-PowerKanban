export interface FlyOutFormResult {
    cancelled: boolean;
    values?: any;
}

export interface FlyOutField {
    type?: string;
    label: string;
    placeholder?: string;
    subtext?: string;
    required?: boolean;
    rows?: number;
    as?: any;
    defaultValue?: string;
}

export interface FlyOutLookupField extends FlyOutField {
    fetchXml: string;
    displayField: string;
    secondaryFields: Array<string>;
    defaultSelectedId: string;
    defaultSelectedName: string;
}

export interface FlyOutForm {
    title: string;
    fields: {[key: string]: FlyOutField };
    resolve: (result: FlyOutFormResult) => void;
    reject: (e: Error) => void;
}
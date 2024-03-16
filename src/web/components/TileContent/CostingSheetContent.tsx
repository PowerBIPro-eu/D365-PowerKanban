import { Badge } from "@fluentui/react-badge";
import { CardHeader } from "@fluentui/react-card";
import {
    Button,
    Field,
    Persona,
    ProgressBar,
} from "@fluentui/react-components";
import { Divider } from "@fluentui/react-divider";
import { Caption1, Subtitle2, Text } from "@fluentui/react-text";
import * as React from "react";
import { mergeClasses } from "@fluentui/react-components";
import {
    CalendarCancel16Regular,
    CalendarLtr20Regular,
    GanttChart24Regular,
    Open16Regular,
} from "@fluentui/react-icons";
import { FC } from "react";

// TODO poradie Inj,Hoses,Pipes,Operating steps(add),Tools

type Props = {
    data: any;
    aditionalData: any;
    openInline: () => void;
    primaryAttriute: string;
};

const CostingSheetContent: FC<Props> = ({
    data,
    aditionalData,
    openInline,
    primaryAttriute,
}) => {
    const [hoses, setHoses] = React.useState("");
    const [pipes, setPipes] = React.useState("");
    const [tools, setTools] = React.useState("");
    const [injectionMaterials, setInjectionMaterials] = React.useState("");
    const [operatingSteps, setOperatingSteps] = React.useState("");
    React.useEffect(() => {
        setHoses(
            `Hoses (${
                aditionalData.hoses.value.find(
                    (a: any) =>
                        a.ddsol_cs_costingsheet_ddsol_cs_costingsheetid ===
                        data[primaryAttriute]
                )?.count ?? 0
            })`
        );
        setPipes(
            `Pipes (${
                aditionalData.pipes.value.find(
                    (a: any) =>
                        a.ddsol_cs_costingsheet_ddsol_cs_costingsheetid ===
                        data[primaryAttriute]
                )?.count ?? 0
            })`
        );
        setTools(
            `Tools (${
                aditionalData.tools.value.find(
                    (a: any) =>
                        a.ddsol_cs_costingsheet_ddsol_cs_costingsheetid ===
                        data[primaryAttriute]
                )?.count ?? 0
            })`
        );
        setInjectionMaterials(
            `Injection Materials (${
                aditionalData.injectionMaterials.value.find(
                    (a: any) =>
                        a.ddsol_cs_costingsheet_ddsol_cs_costingsheetid ===
                        data[primaryAttriute]
                )?.count ?? 0
            })`
        );
        setOperatingSteps(
            `Operating Steps (${
                aditionalData.operatingSteps.value.find(
                    (a: any) =>
                        a.ddsol_cs_costingsheet_ddsol_cs_costingsheetid ===
                        data[primaryAttriute]
                )?.count ?? 0
            })`
        );
    }, []);
    return (
        <>
            <CardHeader
                image={
                    <Badge appearance="filled" color="brand">
                        {data.ddsol_cstitle.slice(0, 2)}
                    </Badge>
                }
                header={
                    <div
                        style={{
                            display: "flex",
                            justifyContent: "space-between",
                        }}
                    >
                        <Subtitle2>
                            <b>{data.ddsol_cstitle}</b>
                        </Subtitle2>
                        <div>
                            <Button
                                icon={<Open16Regular />}
                                size="small"
                                style={{ marginRight: "1rem" }}
                                onClick={openInline}
                            >
                                Open
                            </Button>
                        </div>
                    </div>
                }
            />
            <div
                style={{
                    display: "flex",
                    flexDirection: "row",
                    justifyContent: "space-between",
                }}
            >
                <div
                    style={{
                        display: "flex",
                        flexDirection: "column",
                    }}
                >
                    <div
                        style={{
                            display: "flex",
                            flexDirection: "row",
                            justifyContent: "space-between",
                        }}
                    >
                        <Text>{hoses}</Text>
                        <Text>{pipes}</Text>
                    </div>
                    <Text>{injectionMaterials}</Text>
                </div>
                <div
                    style={{
                        display: "flex",
                        flexDirection: "column",
                    }}
                >
                    <Text>{operatingSteps}</Text>
                    <Text>{tools}</Text>
                </div>
            </div>
        </>
    );
};

export default CostingSheetContent;

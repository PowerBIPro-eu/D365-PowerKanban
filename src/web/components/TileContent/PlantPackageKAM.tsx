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
import { FetchUserAvatar } from "../../domain/FetchUserInfo";
import * as WebApiClient from "xrm-webapi-client";

type Props = {
    styles: any;
    data: any;
    openInline: () => void;
    primaryAttriute: string;
    bipPrices: any;
    costSheets: any;
};

const PlantPackageKam: FC<Props> = ({
    styles,
    data,
    openInline,
    primaryAttriute,
    bipPrices,
    costSheets,
}) => {
    const [avatars, setAvatars] = React.useState<any>({});
    const [customer, setCustomer] = React.useState<any>("");
    const [completions, setCompletions] = React.useState<any>({});

    const discardRecord: (id: string) => void = (id) => {
        WebApiClient.Update({
            entityName: "ddsol_cs_packageplant",
            entityId: id,
            entity: { ddsol_sts_plantpackage: 717170006 },
        });
    };

    const getCustomer = async () => {
        try {
            const project = await WebApiClient.Retrieve({
                entityName: "ddsol_cs_project",
                queryParams: `?$filter=ddsol_cs_projectid eq '${data["_ddsol_project_value"]}'&$expand=ddsol_customer`,
            });
            setCustomer(
                project?.value[0]?.ddsol_customer?.ddsol_customername ?? ""
            );
        } catch (error) {
            console.error("Error retrieving data:", error);
        }
    };

    React.useEffect(() => {
        let purOfficer = FetchUserAvatar(
            data["_ddsol_role_purchasingofficer_value"]
        );
        let costEngineer = FetchUserAvatar(
            data["_ddsol_role_costingengineer_value"]
        );
        console.log("pople: ", purOfficer, costEngineer);
        setAvatars({
            purOfficer: { src: purOfficer },
            costEngineer: { src: costEngineer },
        });

        getCustomer();

        let bips = bipPrices?.value?.filter(
            (bipPrice: any) =>
                bipPrice.ddsol_cs_packageplant_ddsol_cs_packageplantid ==
                data[primaryAttriute]
        );
        let costingSheets = costSheets?.value?.filter(
            (costingSheet: any) =>
                costingSheet.ddsol_cs_packageplant_ddsol_cs_packageplantid ==
                data[primaryAttriute]
        );
        setCompletions({
            bipsTotal:
                bips
                    ?.map((entity: any) => entity.count)
                    ?.reduce((a: any, b: any) => a + b, 0) ?? 0,
            bipsDone:
                bips
                    ?.filter(
                        (entity: any) =>
                            entity?.ddsol_sts_progress === 717170001
                    )
                    ?.map((entity: any) => entity.count)
                    ?.reduce((a: any, b: any) => a + b, 0) ?? 0,
            csTotal:
                costingSheets
                    ?.map((entity: any) => entity.count)
                    ?.reduce((a: any, b: any) => a + b, 0) ?? 0,
            csDone:
                costingSheets
                    ?.filter(
                        (entity: any) =>
                            entity?.ddsol_sts_costingsheet === 717170002 ||
                            entity?.ddsol_sts_costingsheet === 717170004
                    )
                    ?.map((entity: any) => entity.count)
                    ?.reduce((a: any, b: any) => a + b, 0) ?? 0,
        });
    }, []);
    return (
        <>
            <CardHeader
                header={
                    <div
                        style={{
                            display: "flex",
                            flexDirection: "row",
                            justifyContent: "space-between",
                        }}
                    >
                        <Subtitle2>
                            <b>{data.ddsol_name}</b>
                        </Subtitle2>
                        <Badge appearance="tint" color="subtle">
                            {
                                data[
                                    "ddsol_sts_plantpackage@OData.Community.Display.V1.FormattedValue"
                                ]
                            }
                        </Badge>
                    </div>
                }
            />
            <header
                className={mergeClasses(styles.flex)}
                style={{ flexWrap: "wrap" }}
            >
                <Badge appearance="filled" color="brand">
                    {
                        data[
                            "_ddsol_plant_value@OData.Community.Display.V1.FormattedValue"
                        ]
                    }
                </Badge>
                <Badge appearance="filled" color="brand">
                    {customer}
                </Badge>
                <Badge
                    color="brand"
                    shape="rounded"
                    appearance="tint"
                    size="large"
                    icon={<GanttChart24Regular />}
                    title="Project"
                >
                    {
                        data[
                            "_ddsol_project_value@OData.Community.Display.V1.FormattedValue"
                        ]
                    }
                </Badge>
            </header>

            <div
                className="people"
                style={{
                    display: "flex",
                    flexDirection: "row",
                    justifyContent: "space-between",
                    width: "100%",
                }}
            >
                <div style={{ width: "40%" }}>
                    {/* <Caption1>Purchasing Officer</Caption1> */}
                    <Persona
                        name={
                            data[
                                "_ddsol_role_purchasingofficer_value@OData.Community.Display.V1.FormattedValue"
                            ]
                        }
                        secondaryText="Purchasing Officer"
                        avatar={{
                            image: avatars?.purOfficer,
                        }}
                    />
                </div>
                <div style={{ width: "40%" }}>
                    {/* <Caption1>Costing Engineer</Caption1> */}
                    <Persona
                        name={
                            data[
                                "_ddsol_role_costingengineer_value@OData.Community.Display.V1.FormattedValue"
                            ]
                        }
                        secondaryText="Costing Engineer"
                        avatar={{
                            image: avatars?.costEngineer,
                        }}
                    />
                </div>
                {/* <div>
                    {" "}
                    <Text block weight="semibold">
                        Package Plant Status
                    </Text>
                    <Caption1 block className={styles.caption}>
                        {
                            data[
                                "ddsol_sts_plantpackage@OData.Community.Display.V1.FormattedValue"
                            ]
                        }
                    </Caption1>
                </div> */}
            </div>

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
                        width: "100%",
                    }}
                >
                    <div style={{ width: "40%" }}>
                        {" "}
                        <Text block weight="semibold">
                            BIP Status
                        </Text>
                        <Caption1 block className={styles.caption}>
                            {
                                data[
                                    "ddsol_sts_bip@OData.Community.Display.V1.FormattedValue"
                                ]
                            }
                        </Caption1>
                    </div>
                    <div style={{ width: "40%" }}>
                        {" "}
                        <Text block weight="semibold">
                            CS Status
                        </Text>
                        <Caption1 block className={styles.caption}>
                            {
                                data[
                                    "ddsol_sts_costingsheet@OData.Community.Display.V1.FormattedValue"
                                ]
                            }
                        </Caption1>
                    </div>
                </div>
            </div>
            <div
                className={"progresses"}
                style={{
                    display: "flex",
                    flexDirection: "row",
                    justifyContent: "space-between",
                    width: "100%",
                }}
            >
                <Field
                    validationMessage={`BIPS ${completions?.bipsDone}/${completions?.bipsTotal}`}
                    style={{ width: "40%" }}
                    validationState={
                        data["ddsol_sts_bip"] === 717170004
                            ? "success"
                            : data["ddsol_sts_bip"] === 717170002 ||
                              data["ddsol_sts_bip"] === 717170003
                            ? "warning"
                            : "error"
                    }
                >
                    <ProgressBar
                        value={
                            completions?.bipsDone ??
                            0 / completions?.bipsTotal ??
                            0
                        }
                        color={
                            data["ddsol_sts_bip"] === 717170004
                                ? "success"
                                : data["ddsol_sts_bip"] === 717170002 ||
                                  data["ddsol_sts_bip"] === 717170003
                                ? "warning"
                                : "error"
                        }
                    />
                </Field>
                <Field
                    validationMessage={`Costing Sheets ${completions.csDone}/${completions.csTotal}`}
                    style={{ width: "40%" }}
                    validationState={
                        data["ddsol_sts_costingsheet"] >= 717170002 &&
                        data["ddsol_sts_costingsheet"] <= 717170004
                            ? "success"
                            : "error"
                    }
                >
                    <ProgressBar
                        value={completions.csDone / completions.csTotal}
                        color={
                            data["ddsol_sts_costingsheet"] >= 717170002 &&
                            data["ddsol_sts_costingsheet"] <= 717170004
                                ? "success"
                                : "error"
                        }
                    />
                </Field>
            </div>
            <Divider />
            <footer className={mergeClasses(styles.flex, styles.cardFooter)}>
                <div className={styles.flex}>
                    <CalendarLtr20Regular />
                    <Caption1>
                        Created on{" "}
                        <i>{new Date(data.createdon).toDateString()}</i>
                    </Caption1>
                </div>
                <div>
                    {data["ddsol_sts_plantpackage"] === 717170002 ? (
                        <Button
                            icon={<CalendarCancel16Regular />}
                            size="small"
                            style={{ marginRight: "1rem" }}
                            onClick={() => {
                                discardRecord(data[primaryAttriute]);
                            }}
                        >
                            Discard
                        </Button>
                    ) : (
                        <></>
                    )}
                    <Button
                        icon={<Open16Regular />}
                        size="small"
                        style={{ marginRight: "1rem" }}
                        onClick={openInline}
                    >
                        Open
                    </Button>
                </div>
            </footer>
        </>
    );
};

export default PlantPackageKam;

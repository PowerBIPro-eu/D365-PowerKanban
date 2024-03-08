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

type Props = {
    styles: any;
    data: any;
    completions: any;
    openInline: () => void;
};

const PlantPackageCE: FC<Props> = ({
    styles,
    data,
    completions,
    openInline,
}) => {
    return (
        <>
            <CardHeader
                header={
                    <Subtitle2>
                        <b>{data.ddsol_name}</b>
                    </Subtitle2>
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
                    Customer
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

            <div className="people" style={{ display: "flex" }}>
                <div style={{ flex: "50%" }}>
                    <Caption1>Purchasing Officer</Caption1>
                    <Persona
                        name="Kevin Sturgis"
                        secondaryText="Available"
                        presence={{ status: "available" }}
                        avatar={{
                            image: {
                                src: "https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-react-assets/persona-male.png",
                            },
                        }}
                    />
                </div>
                <div style={{ flex: "50%" }}>
                    <Caption1>Costing Engineer</Caption1>
                    <Persona
                        name="Kevin Sturgis"
                        secondaryText="Available"
                        presence={{ status: "available" }}
                        avatar={{
                            image: {
                                src: "https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-react-assets/persona-male.png",
                            },
                        }}
                    />
                </div>
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
                    }}
                >
                    <div>
                        <Text block weight="semibold">
                            Acceptance Status
                        </Text>
                        <Caption1 block className={styles.caption}>
                            {
                                data[
                                    "ddsol_sts_acceptance@OData.Community.Display.V1.FormattedValue"
                                ]
                            }
                        </Caption1>
                    </div>
                    <div>
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
                    </div>
                </div>
                <div
                    style={{
                        display: "flex",
                        flexDirection: "row",
                        justifyContent: "space-between",
                    }}
                >
                    <div>
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
                    <div>
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
                    <div>
                        {" "}
                        <Text block weight="semibold">
                            Cost-Rate Change
                        </Text>
                        <Caption1 block className={styles.caption}>
                            {
                                data[
                                    "ddsol_sts_costratechange@OData.Community.Display.V1.FormattedValue"
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
                }}
            >
                {
                    <Field validationMessage={completions.bip}>
                        <ProgressBar value={3 / 7} color="error" />
                    </Field>
                }

                <Field
                    validationMessage={completions.fs}
                    validationState="warning"
                >
                    <ProgressBar value={5 / 6} color="warning" />
                </Field>

                <Field
                    validationMessage={completions.cs}
                    validationState="success"
                >
                    <ProgressBar value={7 / 7} color="success" />
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
                    <Button
                        icon={<CalendarCancel16Regular />}
                        size="small"
                        style={{ marginRight: "1rem" }}
                        onClick={() => {
                            alert("discarded");
                        }}
                    >
                        Discard
                    </Button>
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

export default PlantPackageCE;

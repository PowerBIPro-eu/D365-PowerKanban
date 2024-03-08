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
    openInline: () => void;
};

const CostingSheetContent: FC<Props> = ({ styles, data, openInline }) => {
    return (
        <>
            <CardHeader
                image={
                    <Badge appearance="filled" color="brand">
                        {data.ddsol_cstitle.slice(0, 2)}
                    </Badge>
                }
                header={
                    <Subtitle2>
                        <b>{data.ddsol_cstitle}</b>
                    </Subtitle2>
                }
            />
            <Divider />
            <footer className={mergeClasses(styles.flex, styles.cardFooter)}>
                <div className={styles.flex}>
                    <CalendarLtr20Regular />
                    <Caption1>
                        Modified on{" "}
                        <i>{new Date(data.modifiedon).toDateString()}</i>
                    </Caption1>
                </div>
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
            </footer>
        </>
    );
};

export default CostingSheetContent;

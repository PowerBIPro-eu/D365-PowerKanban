import { Button } from "@fluentui/react-components";
import { Dialog, DialogFooter, DialogType } from "@fluentui/react/lib/Dialog";
import * as React from "react";

interface UserInputModalProps {
    title: string;
    show: boolean;
    yesCallBack?: (value: string) => void;
    noCallBack?: () => void;
    finally?: () => void;
    okButtonDisabled?: boolean;
}

export class UserInputModal extends React.PureComponent<
    UserInputModalProps,
    undefined
> {
    constructor(props: UserInputModalProps) {
        super(props);

        this.triggerCallback = this.triggerCallback.bind(this);
        this.callIfDefined = this.callIfDefined.bind(this);
        this.setValue = this.setValue.bind(this);
    }

    callIfDefined(callBack: (value?: string) => void) {
        if (callBack) {
            callBack();
        }
    }

    setValue(e: any) {
        const text = e.target.value;

        this.setState({
            value: text,
        });
    }

    triggerCallback(choice: boolean) {
        if (choice) {
            this.callIfDefined(this.props.yesCallBack);
        } else {
            this.callIfDefined(this.props.noCallBack);
        }

        this.callIfDefined(this.props.finally);
    }

    render() {
        const modelProps = {
            isBlocking: true,
        };

        const dialogContentProps = {
            type: DialogType.normal,
            title: this.props.title,
        };

        return (
            <Dialog
                hidden={!this.props.show}
                dialogContentProps={dialogContentProps}
                modalProps={modelProps}
                onDismiss={() => this.triggerCallback(false)}
            >
                {this.props.show && this.props.children}

                <DialogFooter>
                    <Button
                        onClick={() => this.triggerCallback(true)}
                        disabled={this.props.okButtonDisabled}
                    >
                        Ok
                    </Button>
                    <Button onClick={() => this.triggerCallback(false)}>
                        Cancel
                    </Button>
                </DialogFooter>
            </Dialog>
        );
    }
}

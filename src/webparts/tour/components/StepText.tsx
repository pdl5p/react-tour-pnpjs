import * as React from 'react';

import { ICustomCollectionField } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { TextField } from 'office-ui-fabric-react';
import styles from './StepText.module.scss';

export interface IStepTextProps {
    onUpdate: (fieldId: string, value: any) => void;
    onError: (fieldId: string, error: string) => void;
    value: string;
    field: ICustomCollectionField;
}

export const StepText = (props: IStepTextProps) => {
    const { onUpdate, value, field, onError } = props;

    const onChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        if (onUpdate) {
            onUpdate(field.id, newValue);
            if (onError) {
                if (newValue) {
                    onError(field.id, "");
                }
                else {
                    onError(field.id, "Please enter the text for this step");
                }
            }
        }
    };

    React.useEffect(() => {

        if(!value) {
            onError(field.id, "Please enter the text for this step");
        }

    }, []);

    return (
        <div className={styles.container}>
        <TextField
            value={value}
            onChange={onChange}
            multiline={true}
            rows={4}
        />
        </div>
    );
};
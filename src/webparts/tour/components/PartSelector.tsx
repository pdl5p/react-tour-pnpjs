import * as React from 'react';

import { Dropdown, TextField, IDropdownOption } from 'office-ui-fabric-react';
import styles from './PartSelector.module.scss';
import { ICustomCollectionField } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface IPartSelectorProps {
    options: IDropdownOption[];
    onUpdate: (fieldId: string, value: any) => void;
    onError: (fieldId: string, error: string) => void;
    value: string;
    field: ICustomCollectionField;
}

const { useState, useEffect } = React;

export const PartSelector = (props: IPartSelectorProps) => {

    const {options, onUpdate, value, field, onError } = props;

    const [dropdownValue, setDropdownValue] = useState("");
    const [textValue, setTextValue] = useState("");

    useEffect(() => {
        
        if(options.filter(o => o.key === value).length > 0) {
            setDropdownValue(value);
        }
        else if(value) {
            setDropdownValue("custom");
            setTextValue(value);
        }
        else{
            onError(field.id, "Please select a web part or enter a custom selector");
        }
            
        return () => {
            // cleanup
        };
    }, []);

    const onDropdownChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
        const key = option.key as string;
        
        setDropdownValue(key);
 
        if(key === 'custom') 
        {
            if(textValue){
                onUpdate(field.id, textValue);
                onError(field.id, "");
            }
            else{
                onUpdate(field.id, "");
                onError(field.id, "Please enter a valid CSS selector");
            }
        }else{
            onUpdate(field.id, key);
            onError(field.id, "");
        }
    };

    const onTextFieldChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        setTextValue(newValue);
    };

    const onTextBlur = (event: React.FocusEvent<HTMLInputElement | HTMLTextAreaElement>) => {
        
        onUpdate(field.id, textValue);
        if(!textValue) {
            onError(field.id, "Please enter a valid CSS selector");
        }
        else{
            onError(field.id, "");
        }
    };

    const isCustom = dropdownValue === 'custom';

    return (<div className={styles.container}>
        <Dropdown label="" options={options} onChange={onDropdownChange} selectedKey={dropdownValue} />
        {isCustom && <TextField label="" placeholder="Enter custom selector" 
            value={textValue} 
            onChange={onTextFieldChange} 
            onBlur={onTextBlur}
        />}
    </div>);
};
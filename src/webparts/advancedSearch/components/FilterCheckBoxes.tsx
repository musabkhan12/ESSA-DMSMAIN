import React, { useState } from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import { SearchAggregation } from '@microsoft/microsoft-graph-types';

export interface IFilterCheckProps
{
    refiner:SearchAggregation ;
    handleCheckboxChange?:any;
    searchRefinerFilters?:Map<string,string[]>;
}

export const FilterCheckBox = (props:IFilterCheckProps) => {
    return (
        <div>
        { props.refiner.buckets.map(value => (
                <div className="form-check newcheck" key={value.key}>
                    
                    <input
                        type="checkbox"
                        id={`${props.refiner.field}-${value.key}`}
                        data-field={props.refiner.field}
                        data-token={`${value.aggregationFilterToken}`}
                        checked={(Array.from(props.searchRefinerFilters.keys()).includes(props.refiner.field))?props.searchRefinerFilters.get(props.refiner.field).includes(value.aggregationFilterToken):false}
                        //checked={Array.from(searchRefinerFilters.keys()).includes(refiner.field)}
                        className="form-check-input"
                        onChange={(e) => props.handleCheckboxChange(props.refiner.field, value.aggregationFilterToken, e.target.checked)}
                    />
                    <label htmlFor={`${props.refiner.field}-${value.key}`} className="form-check-label">{value.key}</label>
                </div>
            ))}
        </div>

    )
}
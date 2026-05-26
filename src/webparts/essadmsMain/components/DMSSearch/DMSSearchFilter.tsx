// FilterComponent.tsx
import React, { useState } from "react";

export interface FilterComponentProps {
  properties: string[]  
  value: any;
  onSubmit: (property: string, equalityop:string, value: any) => void;
}
export enum PropertyValueType
{
    Text,Datetime,Number
}
export interface PropertyType
{
   Key:string;
   Value:string;
   ValueType:PropertyValueType; 
}
export  const DMSSearchFilter: React.FC<FilterComponentProps> = ({ properties, value, onSubmit }) => {
  const [inputValue, setInputValue] = useState(value);
  const [selectedProperty, setSelectedProperty] = useState("Title");
  const [comparisonOperator, setComparisonOperator] = useState("EqualTo");

  const handleSubmit = () => {
    onSubmit(selectedProperty,comparisonOperator, inputValue);
  };

  const handlePropertySelect = (event: React.ChangeEvent<HTMLSelectElement>) => {
    setSelectedProperty(event.target.value);
  };

  const handleInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setInputValue(event.target.value);
  };

  const handleComparisonOperatorChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    setComparisonOperator(event.target.value);
  };

  return (
      <>
          <div className="col-xl-3">
              <select value={selectedProperty} onChange={handlePropertySelect} className="form-control">
                  {properties.map((property) => (
                      <option key={property} value={property}>
                          {property}
                      </option>
                  ))}
              </select>
          </div>
          <div className="col-xl-3">
              <select value={comparisonOperator} onChange={handleComparisonOperatorChange} className="form-control">
                  <option value="EqualTo">Equal To</option>
                  <option value="Contains">Contains</option>
                  <option value="GreaterThan">Greater Than</option>
                  <option value="LessThan">Less Than</option>
                  <option value="GreaterThanOrEqualTo">Greater Than or Equal To</option>
                  <option value="LessThanOrEqualTo">Less Than or Equal To</option>
              </select>
          </div>
          <div className="col-xl-4">
              {selectedProperty === "Created" || selectedProperty === "Modified" ? (
                  <>
                      <input type="date" value={inputValue.from} onChange={(event) => setInputValue({ ...inputValue, from: event.target.value })} className="form-control"/>
                      <input type="date" value={inputValue.to} onChange={(event) => setInputValue({ ...inputValue, to: event.target.value })} className="form-control"/>
                  </>
              ) : selectedProperty === "Title" ? (
                  <input type="text" value={inputValue} onChange={handleInputChange} className="form-control"/>
              ) : null}
          </div>
          <div className="col-xl-1">
              <button type="button" onClick={handleSubmit}>Submit</button>
          </div>
      </>
  );
};



export default DMSSearchFilter;
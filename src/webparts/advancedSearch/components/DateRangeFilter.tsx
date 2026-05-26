import React, { useState } from 'react';

export interface IDateRangeFilterProps {
    onFilterChange: (filter: string) => void; // Callback to pass the filter query
}

export const DateRangeFilter: React.FC<IDateRangeFilterProps> = ({ onFilterChange }) => {
    const [startDate, setStartDate] = useState<string>('');
    const [endDate, setEndDate] = useState<string>('');

     // Generate filter query based on input
     const generateFilter = (): string => {
        if (startDate && endDate) {
            // return `LastModifiedTime:Range(${startDate}..${endDate})`;
            return `LastModifiedTime>=${startDate} AND LastModifiedTime<=${endDate}`;
        } else if (startDate) {
            return `LastModifiedTime>=${startDate}`;
        } else if (endDate) {
            return `LastModifiedTime<=${endDate}`;
        }
        return '';
    };
    
    // Update filter query on date change
    const handleFilterChange = () => {
        // let filter = '';
        // if (startDate && endDate) {
        //     filter = `LastModifiedTime:Range(${startDate}..${endDate})`;
        // } else if (startDate) {
        //     filter = `LastModifiedTime>=${startDate}`;
        // } else if (endDate) {
        //     filter = `LastModifiedTime<=${endDate}`;
        // }

        // onFilterChange(filter); // Pass filter to parent
    };

    // Handle Apply button click
    const handleApply = () => {
        const filter = generateFilter();
        onFilterChange(filter); // Pass the filter query
    };


    return (
        <div  className="mt-3 p-3 pt-0">
            
            <div style={{background:"#fff",border:'1px solid #1fb0e5', borderRadius:'20px'}} className="row g-3">
            <div className="col-md-3">
            <h5 style={{fontWeight:'600'}} className='mb-2 mt-2 font-16 text-center'>Filter by Published Date</h5>
                </div>
                <div className='col-md-1'><label htmlFor="startDate" className="form-label mb-0 mt-2">Start Date</label></div>
                <div className="col-md-2">
                    
                    <input
                        type="date"
                        id="startDate"
                        className="form-control"
                        value={startDate}
                        onChange={(e) => {
                            setStartDate(e.target.value);
                            handleFilterChange();
                        }}
                    />
                </div>
                <div className="col-md-1">  <label htmlFor="endDate" className="form-label mb-0 mt-2">End Date</label> </div>
                <div className="col-md-2">
                   
                    <input
                        type="date"
                        id="endDate"
                        className="form-control"
                        value={endDate}
                        onChange={(e) => {
                            setEndDate(e.target.value);
                            handleFilterChange();
                        }}
                    />
                </div>
                <div className="col-md-2 mt-3 mb-3">
                    <button type='button' style={{marginTop:'0rem',fontSize:'1rem',padding:'6px 20px', minWidth:'auto'}} className="btn btn-primary" onClick={handleApply}>
                        Apply
                    </button>
                </div>
            </div>
        </div>
    );
};


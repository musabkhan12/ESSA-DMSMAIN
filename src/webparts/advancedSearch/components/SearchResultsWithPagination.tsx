import React, { useState } from 'react';
import { Pagination, Form, Container } from 'react-bootstrap';
import TagsComponent from './SearchResultsTags';
import { GetFieldName, GetFieldValue,enumfieldtype } from './Common';
import { IField } from './AdvancedSearch';
// import TagsComponent from './'; // Adjust the import path as necessary

export interface SearchResult {
    Title: string;
    Summary: string;
    Properties: { [key: string]: string };
    Path:string;
}

export interface SearchResultsWithPaginationProps {
    searchResult: SearchResult[];
    fieldnamesmapping: { [key: string]: string };
    fieldtypemappings:IField[];
}

export const SearchResultsWithPagination: React.FC<SearchResultsWithPaginationProps> = ({ searchResult, fieldnamesmapping,fieldtypemappings }) => {
    const [currentPage, setCurrentPage] = useState(1);
    const [itemsPerPage, setItemsPerPage] = useState(10);

    const handlePageChange = (pageNumber: number) => {
        setCurrentPage(pageNumber);
    };

    const handleItemsPerPageChange = (event: React.ChangeEvent<any>) => {
        setItemsPerPage(Number(event.target.value));
        setCurrentPage(1); // Reset to first page when items per page changes
    };

    const indexOfLastItem = currentPage * itemsPerPage;
    const indexOfFirstItem = indexOfLastItem - itemsPerPage;
    const currentItems = searchResult.slice(indexOfFirstItem, indexOfLastItem);

    const totalPages = Math.ceil(searchResult.length / itemsPerPage);

    const renderPaginationItems = () => {
        const pageItems = [];
        const maxPageItems = 5; // Maximum number of page items to display

        if (totalPages <= maxPageItems) {
            for (let i = 1; i <= totalPages; i++) {
                pageItems.push(
                    <Pagination.Item key={i} active={i === currentPage} onClick={() => handlePageChange(i)}>
                        {i}
                    </Pagination.Item>
                );
            }
        } else {
            let startPage = Math.max(1, currentPage - Math.floor(maxPageItems / 2));
            let endPage = startPage + maxPageItems - 1;

            if (endPage > totalPages) {
                endPage = totalPages;
                startPage = endPage - maxPageItems + 1;
            }

            if (startPage > 1) {
                pageItems.push(
                    <Pagination.Item key={1} onClick={() => handlePageChange(1)}>
                        1
                    </Pagination.Item>
                );
                if (startPage > 2) {
                    pageItems.push(<Pagination.Ellipsis key="start-ellipsis" />);
                }
            }

            for (let i = startPage; i <= endPage; i++) {
                pageItems.push(
                    <Pagination.Item key={i} active={i === currentPage} onClick={() => handlePageChange(i)}>
                        {i}
                    </Pagination.Item>
                );
            }

            if (endPage < totalPages) {
                if (endPage < totalPages - 1) {
                    pageItems.push(<Pagination.Ellipsis key="end-ellipsis" />);
                }
                pageItems.push(
                    <Pagination.Item key={totalPages} onClick={() => handlePageChange(totalPages)}>
                        {totalPages}
                    </Pagination.Item>
                );
            }
        }

        return pageItems;
    };

    return (
        <div>
           
            {/* <div className="pagination-container">
                <Pagination className="flex-wrap">
                    <Pagination.First onClick={() => handlePageChange(1)} disabled={currentPage === 1} />
                    <Pagination.Prev onClick={() => handlePageChange(currentPage - 1)} disabled={currentPage === 1} />
                    {renderPaginationItems()}
                    <Pagination.Next onClick={() => handlePageChange(currentPage + 1)} disabled={currentPage === totalPages} />
                    <Pagination.Last onClick={() => handlePageChange(totalPages)} disabled={currentPage === totalPages} />
                </Pagination>
            </div> */}
            <div className="row p-2 pt-0">
                {currentItems.map((res, index) => (
                    <div key={index} className="col-sm-12 mt-0 mb-2 search-result">
                        <div style={{borderRadius:'5px'}}  className="card h-100">
                            <div className="card-body">
                                <a href={res.Path} target='_blank'><h5   className="card-title text-dark hover fw-bold">{res.Title}</h5></a>
                                {/* <p className="card-text">{res.Summary}</p> */}
                                <p style={{fontSize:'14px', color:'#666'}} className="card-text font-14"><span dangerouslySetInnerHTML={{ __html:`${res.Summary.replace(/<c0>/g, "<strong>").replace(/<\/c0>/g, "</strong>")}` }} /></p>
                            </div>
                            {/* <TagsComponent tags={Object.entries(res.Properties).map(([key, value]) => `${fieldnamesmapping[key] ? fieldnamesmapping[key] : key}: ${value}`)} /> */}
                            <TagsComponent tags={Object.entries(res.Properties).map(([key, value]) => `${GetFieldName(key)}: ${GetFieldValue(fieldtypemappings,key,value)}`)} />
                        </div>
                    </div>
                ))}
            </div>
            <div className='row'>
                <div className='col-sm-4'>
                <div className="d-flex justify-content-start align-items-center mt-0 mb-2">
                <Form.Group controlId="itemsPerPageSelect" className='newdesign'>
                    <Form.Label>Items per page:</Form.Label>
                    <Form.Control as="select" value={itemsPerPage} onChange={handleItemsPerPageChange}>
                        <option value={10}>10</option>
                        <option value={20}>20</option>
                        <option value={30}>30</option>
                        <option value={50}>50</option>
                    </Form.Control>
                </Form.Group>
            </div>
                </div>
                <div className='col-sm-8'>
                <div className="pagination-container">
                <Pagination className="flex-wrap">
                    <Pagination.First onClick={() => handlePageChange(1)} disabled={currentPage === 1} />
                    <Pagination.Prev onClick={() => handlePageChange(currentPage - 1)} disabled={currentPage === 1} />
                    {renderPaginationItems()}
                    <Pagination.Next onClick={() => handlePageChange(currentPage + 1)} disabled={currentPage === totalPages} />
                    <Pagination.Last onClick={() => handlePageChange(totalPages)} disabled={currentPage === totalPages} />
                </Pagination>
            </div>

                    </div>
            </div>
            
            <style>jsx{`
                .pagination-container {
                    display: flex;
                    justify-content: center;
                    overflow-x: auto;
                }
                .pagination-container .pagination {
                    flex-wrap: nowrap;
                }
            `}</style>
        </div>
    );
};


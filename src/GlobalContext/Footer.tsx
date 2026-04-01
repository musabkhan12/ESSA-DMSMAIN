import * as React from 'react';
 
export interface  IFooterProps {
   
  }
  export const Footer:React.FunctionComponent<IFooterProps> = (
    props:React.PropsWithChildren<IFooterProps>
  ) => {
console.log();
 
    return (
        <footer className="footer">
        <div className="container-fluid">
        <div className="row">
        <div className="col-md-12">
        {/* <div>{new Date().getFullYear()} © alrostamanigroup </div> */}
        <div className='font-14'>2024 © alrostamanigroup </div>
        </div>
        <div className="col-md-6">
        {/* <div className="d-none d-md-flex gap-4 align-item-center justify-content-md-end footer-links">
        <a href="javascript: void(0);">About</a>
        <a href="javascript: void(0);">Support</a>
        <a href="javascript: void(0);">Contact Us</a>
        </div> */}
        </div>
        </div>
        </div>
        </footer>
    )
  }
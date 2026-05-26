import React, { useState } from "react";
import "bootstrap/dist/css/bootstrap.min.css";

export interface ITagsProps
{
    tags:string[];
}

const TagsComponent = (props:ITagsProps) => {
  const [tags, setTags] = useState(["React", "JavaScript", "CSS"]);
  const [inputValue, setInputValue] = useState("");

  // Handle adding a new tag
//   const addTag = () => {
//     if (inputValue.trim() && !tags.includes(inputValue.trim())) {
//       setTags([...tags, inputValue.trim()]);
//       setInputValue("");
//     }
//   };

//   // Handle removing a tag
//   const removeTag = (tagToRemove:any) => {
//     setTags(tags.filter((tag) => tag !== tagToRemove));
//   };

  return (
    <div className="container my-0 newbutton">
      <div className="d-flex flex-wrap gap-2 mb-3">
        {props.tags.map((tag, index) => (
          <span style={{fontSize:'12px'}}
            key={index}
            className="badge bg-primary d-flex align-items-center"
          >
            {tag}
            {/* <button
              type="button"
              className="btn-close btn-close-white ms-2"
              aria-label="Remove"
              onClick={() => removeTag(tag)}
              style={{ fontSize: "0.7rem" }}
            ></button> */}
          </span>
        ))}
      </div>
      {/* <div className="input-group">
        <input
          type="text"
          className="form-control"
          placeholder="Add a tag"
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
        />
        <button className="btn btn-primary" onClick={addTag}>
          Add
        </button>
      </div> */}
    </div>
  );
};

export default TagsComponent;

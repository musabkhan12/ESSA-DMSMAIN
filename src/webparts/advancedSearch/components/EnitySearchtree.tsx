import React, { useState } from "react";
import "bootstrap/dist/css/bootstrap.min.css";


// import "bootstrap-icons/font/bootstrap-icons.css";

interface TreeNodeData {
  id: number;
  label: string;
  checked: boolean;
  expanded?: boolean;
  children?: TreeNodeData[];
}

interface TreeNodeProps {
  node: TreeNodeData;
  onCheck: (id: number) => void;
  onExpand: (id: number) => void;
}

const TreeNode: React.FC<TreeNodeProps> = ({ node, onCheck, onExpand }) => {
  const hasChildren = node.children && node.children.length > 0;

  return (
    <div className="ms-3">
      <div className="d-flex align-items-center">
        {hasChildren && (
          <span
            onClick={() => onExpand(node.id)}
            className={`me-2 ${hasChildren ? "cursor-pointer" : ""}`}
            style={{ userSelect: "none", cursor: "pointer" }}
          >
            {node.expanded ? (
              <i className="bi bi-caret-down-fill"></i>
            ) : (
              <i className="bi bi-caret-right-fill"></i>
            )}
          </span>
        )}
        <input
          type="checkbox"
          className="form-check-input me-2"
          checked={node.checked}
          onChange={() => onCheck(node.id)}
        />
        <span>{node.label}</span>
      </div>
      {hasChildren && node.expanded && (
        <div>
          {node.children?.map((child) => (
            <TreeNode
              key={child.id}
              node={child}
              onCheck={onCheck}
              onExpand={onExpand}
            />
          ))}
        </div>
      )}
    </div>
  );
};

const TreeView: React.FC = () => {
  const [treeData, setTreeData] = useState<TreeNodeData[]>([
    {
      id: 1,
      label: "Parent 1",
      checked: false,
      expanded: false,
      children: [
        {
          id: 2,
          label: "Child 1.1",
          checked: false,
          children: [],
        },
        {
          id: 3,
          label: "Child 1.2",
          checked: false,
          expanded: false,
          children: [
            {
              id: 4,
              label: "Sub-child 1.2.1",
              checked: false,
            },
          ],
        },
      ],
    },
    {
      id: 5,
      label: "Parent 2",
      checked: false,
      expanded: false,
      children: [],
    },
  ]);

  const handleCheck = (id: number): void => {
    const toggleCheck = (nodes: TreeNodeData[]): TreeNodeData[] =>
      nodes.map((node) => ({
        ...node,
        checked: node.id === id ? !node.checked : node.checked,
        children: node.children ? toggleCheck(node.children) : node.children,
      }));

    setTreeData(toggleCheck(treeData));
  };

  const handleExpand = (id: number): void => {
    const toggleExpand = (nodes: TreeNodeData[]): TreeNodeData[] =>
      nodes.map((node) => ({
        ...node,
        expanded: node.id === id ? !node.expanded : node.expanded,
        children: node.children ? toggleExpand(node.children) : node.children,
      }));

    setTreeData(toggleExpand(treeData));
  };

  return (
    <div>
      {treeData.map((node) => (
        <TreeNode
          key={node.id}
          node={node}
          onCheck={handleCheck}
          onExpand={handleExpand}
        />
      ))}
    </div>
  );
};

export default TreeView;

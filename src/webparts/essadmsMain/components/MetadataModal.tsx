import * as React from "react";
import { useState, useEffect } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { Modal } from "react-bootstrap";


interface FieldInfo {
    InternalName: string;
}

interface ColumnData {
    ColumnName: string;
    ColumnType: string;
}

interface ExistingColumn {
    ID: number;
    Title?: string;
    ColumnName?: string;
    ColumnType?: string;
    IsRequired?: boolean;
    SiteName?: string;
    DocumentLibraryName?: string;
    // Aman 25/3/26
    IsRename?: string;
}

interface SiteInfo {
    siteCollection: string;
    subsiteName: string;
    documentLibraryName: string;
    dmsListUrl?: string;
}

const MetadataModal: React.FC<{
    show: boolean;
    onClose: () => void;
    file: any;
    context: any;
}> = ({ show, onClose, file, context }) => {
    const [existingColumns, setExistingColumns] = useState<ExistingColumn[]>([]);
    const [newColumns, setNewColumns] = useState<ColumnData[]>([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState<string>("");
    const [availableFields, setAvailableFields] = useState<string[]>([]);

    /* ================= GET SITE INFO FROM FILE OBJECT ================= */

    const getSiteInfo = (): SiteInfo | null => {
        try {
            console.log("🔍 Getting site info from file object...");
            console.log("📁 File Object:", file);

            const folderPath = file?.FolderPath || file?.ServerRelativeUrl || file?.FileRef;

            if (folderPath) {
                console.log("📍 Folder Path:", folderPath);

                const pathParts = folderPath.split('/').filter((p: string) => p && p.trim());

                let siteCollection = "";
                let subsiteName = "";

                const sitesIndex = pathParts.findIndex((p: string) => p.toLowerCase() === 'sites');

                if (sitesIndex !== -1 && pathParts.length > sitesIndex + 1) {
                    siteCollection = pathParts[sitesIndex + 1];

                    if (pathParts.length > sitesIndex + 2) {
                        subsiteName = pathParts[sitesIndex + 2];
                    } else {
                        subsiteName = siteCollection;
                    }
                }

                console.log("✅ From FolderPath - Site Collection:", siteCollection, "Subsite:", subsiteName);

                if (siteCollection) {
                    const documentLibraryName =
                        file?.DocumentLibraryName ||
                        file?.FolderName ||
                        file?.__libraryTitle ||
                        file?.LibraryName ||
                        file?.listTitle ||
                        "";

                    return {
                        siteCollection,
                        subsiteName,
                        documentLibraryName,
                    };
                }
            }

            const siteTitle = file?.SiteTitle || file?.__siteTitle;
            const siteUrl = file?.__siteUrl || file?.siteUrl;

            if (siteTitle && siteUrl) {
                console.log("📍 Site Title:", siteTitle);
                console.log("📍 Site URL:", siteUrl);

                const urlParts = siteUrl.split('/sites/');
                let siteCollection = "";

                if (urlParts[1]) {
                    const pathParts = urlParts[1].split('/');
                    siteCollection = pathParts[0];
                }

                console.log("✅ From SiteTitle/URL - Site Collection:", siteCollection, "Subsite:", siteTitle);

                if (siteCollection) {
                    const documentLibraryName =
                        file?.DocumentLibraryName ||
                        file?.FolderName ||
                        file?.__libraryTitle ||
                        file?.LibraryName ||
                        file?.listTitle ||
                        "";

                    return {
                        siteCollection,
                        subsiteName: siteTitle,
                        documentLibraryName,
                    };
                }
            }

            const currentWebUrl = context.pageContext.web.absoluteUrl;
            console.log("⚠️ Falling back to current context:", currentWebUrl);

            const urlParts = currentWebUrl.split('/sites/');
            if (!urlParts[1]) {
                console.error("❌ Cannot parse site URL");
                return null;
            }

            const pathParts = urlParts[1].split('/');
            const siteCollection = pathParts[0];
            const subsiteParts = pathParts.slice(1).filter((p: string) => p && p.trim());
            const subsiteName = subsiteParts.length > 0
                ? subsiteParts[subsiteParts.length - 1]
                : siteCollection;

            const documentLibraryName =
                file?.DocumentLibraryName ||
                file?.FolderName ||
                file?.__libraryTitle ||
                file?.LibraryName ||
                file?.listTitle ||
                "";

            console.log("✅ From Context - Site Collection:", siteCollection, "Subsite:", subsiteName);

            return {
                siteCollection,
                subsiteName,
                documentLibraryName,
            };

        } catch (err) {
            console.error("❌ Error in getSiteInfo:", err);
            return null;
        }
    };

    /* ================= GET DMSPreviewFormMaster URL ================= */

    const getDMSPreviewFormMasterUrl = (siteCollection: string): string => {
        const baseUrl = context.pageContext.web.absoluteUrl.split('/sites/')[0];
        return `${baseUrl}/sites/${siteCollection}`;
    };

    /* ================= FETCH FIELDS ================= */

    const fetchAvailableFields = async (siteInfo: SiteInfo): Promise<string[]> => {
        try {
            const dmsUrl = getDMSPreviewFormMasterUrl(siteInfo.siteCollection);
            siteInfo.dmsListUrl = dmsUrl;

            console.log(`🔍 Fetching fields from: ${dmsUrl}/DMSPreviewFormMaster`);

            const sp = spfi(dmsUrl).using(SPFx(context));

            const fields: FieldInfo[] = await sp.web.lists
                .getByTitle("DMSPreviewFormMaster")
                .fields
                .select("InternalName")();

            const names = fields.map((f: FieldInfo) => f.InternalName);
            console.log("✅ Available fields:", names);

            setAvailableFields(names);

            return names;
        } catch (err) {
            console.error("❌ Error fetching fields:", err);
            return [];
        }
    };

    // Aman 25/3/26 - Build select query based on available fields
    // const buildSelectQuery = (fields: string[]): string => {
    //     const base = ["ID", "Title"];
    //     const optional = [
    //         "SiteName",
    //         "DocumentLibraryName",
    //         "ColumnName",
    //         "ColumnType",
    //         "IsRequired",
    //         "Sequence",
    //         "AddorRemoveThisColumn",
    //         "IsDocumentLibrary",
    //         "IsPrivate",
    //         "Modified",
    //     ];

    //     return base.concat(
    //         optional.filter(o =>
    //             fields.some(f => f.toLowerCase() === o.toLowerCase())
    //         )
    //     ).join(",");
    // };

    // Aman 25/3/26 
    const buildSelectQuery = (fields: string[]): string => {
        const base = ["ID", "Title"];
        const optional = [
            "SiteName",
            "DocumentLibraryName",
            "ColumnName",
            "IsRename",            //Aman 24/3/26
            "ColumnType",
            "IsRequired",
            "Sequence",
            "AddorRemoveThisColumn",
            "IsDocumentLibrary",
            "IsPrivate",
            "Modified",
        ];
 
        // Map through optional fields and find the REAL InternalName from the list
        const matchedOptional = optional.map(o =>
            fields.find(f => f.toLowerCase() === o.toLowerCase())
        ).filter(f => f !== undefined) as string[];
 
        return base.concat(matchedOptional).join(",");
    };

    /* ================= FETCH EXISTING COLUMNS ================= */

    const fetchExistingColumns = async () => {
        setLoading(true);
        setError("");

        try {
            const siteInfo = getSiteInfo();
            if (!siteInfo) {
                throw new Error("Cannot determine site context");
            }

            if (!siteInfo.documentLibraryName) {
                throw new Error("Document library name not found");
            }

            console.log("🔍 Fetching existing columns for:", siteInfo);

            const fields = await fetchAvailableFields(siteInfo);
            const selectQuery = buildSelectQuery(fields);

            console.log("📝 Select query:", selectQuery);

            const listUrl = siteInfo.dmsListUrl;
            if (!listUrl) {
                throw new Error("DMSPreviewFormMaster location not found");
            }

            console.log("📍 Fetching from:", listUrl);

            const sp = spfi(listUrl).using(SPFx(context));

            let filter = `DocumentLibraryName eq '${siteInfo.documentLibraryName}'`;

            if (fields.some(f => f.toLowerCase() === "sitename")) {
                filter += ` and SiteName eq '${siteInfo.subsiteName}'`;
            }

            console.log("🔎 Filter:", filter);

            const items: ExistingColumn[] = await sp.web.lists
                .getByTitle("DMSPreviewFormMaster")
                .items
                .select(selectQuery)
                .filter(filter)
                .orderBy("Modified", false)
                .top(5000)();

            console.log("✅ Fetched existing columns:", items);
            setExistingColumns(items);

        } catch (err: any) {
            console.error("❌ Error fetching columns:", err);
            setError(err.message || "Failed to fetch columns");
            setExistingColumns([]);
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        if (show && file) {
            fetchExistingColumns();
        }
    }, [show, file]);

    /* ================= ADD/REMOVE COLUMNS ================= */

    const handleAddNewColumn = (e?: React.MouseEvent<HTMLButtonElement>) => {
        if (e) {
            e.preventDefault();
            e.stopPropagation();
        }

        setNewColumns([
            ...newColumns,
            { ColumnName: "", ColumnType: "Single Line of Text" }
        ]);
    };

    const handleNewColumnChange = (index: number, key: string, value: string) => {
        const copy = [...newColumns];
        copy[index] = { ...copy[index], [key]: value };
        setNewColumns(copy);
    };

    const handleRemoveColumn = (index: number, e?: React.MouseEvent<HTMLButtonElement>) => {
        if (e) {
            e.preventDefault();
            e.stopPropagation();
        }

        setNewColumns(newColumns.filter((_, idx) => idx !== index));
    };

    /* ================= SAVE NEW COLUMNS ================= */

    const handleSubmit = async (e?: React.MouseEvent<HTMLButtonElement>) => {
        if (e) {
            e.preventDefault();
            e.stopPropagation();
        }

        const valid = newColumns.filter(c => c.ColumnName.trim());
        if (!valid.length) {
            alert("⚠️ Please add at least one field!");
            return;
        }

        setLoading(true);
        setError("");

        try {
            const siteInfo = getSiteInfo();
            if (!siteInfo) {
                throw new Error("Cannot determine site context");
            }

            console.log("💾 Saving new columns...");
            console.log("📍 Site Collection:", siteInfo.siteCollection);
            console.log("🏷️ Subsite Name:", siteInfo.subsiteName);
            console.log("📁 Library:", siteInfo.documentLibraryName);

            const listUrl = siteInfo.dmsListUrl || getDMSPreviewFormMasterUrl(siteInfo.siteCollection);

            console.log("📍 Saving to:", listUrl);

            const sp = spfi(listUrl).using(SPFx(context));

            for (const col of valid) {
                const payload: any = {
                    Title: col.ColumnName.trim(),
                    ColumnName: col.ColumnName.trim(),
                    ColumnType: col.ColumnType,
                    DocumentLibraryName: siteInfo.documentLibraryName,
                    IsRequired: true,
                    AddorRemoveThisColumn: "Add To Library",
                    IsDocumentLibrary: false,
                     // srs 19/2/26
                    IsInProgress: true, 
                };

                if (availableFields.some(f => f.toLowerCase() === "sitename")) {
                    payload.SiteName = siteInfo.subsiteName;
                }

                console.log("📝 Saving column:", payload);

                await sp.web.lists
                    .getByTitle("DMSPreviewFormMaster")
                    .items
                    .add(payload);
            }

            console.log("✅ All columns saved!");
            alert(`✅ Successfully added ${valid.length} field(s) to ${siteInfo.subsiteName}/${siteInfo.documentLibraryName}!`);

            setNewColumns([]);
            await fetchExistingColumns();

        } catch (err: any) {
            console.error("❌ Error saving columns:", err);
            setError(err.message || "Failed to save columns");
            alert(`❌ Error: ${err.message || "Failed to save columns"}`);
        } finally {
            setLoading(false);
        }
    };

    if (!show) return null;

    /* ================= UI ================= */

    const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>) => {
        if (e.target === e.currentTarget) {
            onClose();
        }
    };

    return (
    <Modal
        show={show}
        onHide={onClose}
        // backdrop="static"
        // ritik
        backdrop={true}
        container={() => document.getElementById('filelistcontainer')}
        centered
        size="lg"
    >
        <Modal.Body style={{ padding: 0 }}>
            <div
                style={{
                    background: "#fff",
                    width: "100%",
                    maxHeight: "85vh",
                    display: "flex",
                    flexDirection: "column",
                    borderRadius: "8px",
                    overflow: "hidden",
                }}
            >
                {/* HEADER */}
                <div
                    style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        padding: "16px 20px",
                        borderBottom: "1px solid #e5e5e5",
                        background: "#fff",
                    }}
                >
                    <h3 style={{ margin: 0, fontSize: "18px", fontWeight: 600, color: "#333" }}>
                        Add Meta Columns
                    </h3>
                    <button
                        onClick={onClose}
                        style={{
                            background: "transparent",
                            border: "2px solid #dc3545",
                            borderRadius: "50%",
                            color: "#dc3545",
                            fontSize: "20px",
                            cursor: "pointer",
                            width: "32px",
                            height: "32px",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            padding: 0,
                            fontWeight: "bold",
                        }}
                        title="Close"
                    >
                        ×
                    </button>
                </div>

                {/* BODY */}
                <div
                    style={{
                        padding: "20px",
                        overflowY: "auto",
                        flex: 1,
                        background: "#fafafa",
                    }}
                >
                    {loading && (
                        <div style={{ textAlign: "center", padding: "40px", color: "#666" }}>
                            <div style={{ fontSize: "32px", marginBottom: "10px" }}>⏳</div>
                            <div>Loading...</div>
                        </div>
                    )}

                    {error && (
                        <div
                            style={{
                                background: "#f8d7da",
                                color: "#721c24",
                                padding: "12px 16px",
                                borderRadius: "4px",
                                marginBottom: "20px",
                                border: "1px solid #f5c6cb",
                                fontSize: "14px",
                            }}
                        >
                            ❌ {error}
                        </div>
                    )}

                    {!loading && !error && (
                        <>
                            {/* EXISTING COLUMNS TABLE */}
                            {existingColumns.length > 0 && (
                                <div style={{ marginBottom: "24px" }}>
                                    <div
                                        style={{
                                            fontSize: "14px",
                                            fontWeight: 600,
                                            color: "#666",
                                            marginBottom: "8px",
                                            padding: "0 4px",
                                        }}
                                    >
                                        Existing Columns ({existingColumns.length})
                                    </div>
                                    <div style={{ background: "#fff", borderRadius: "4px", overflow: "hidden", border: "1px solid #e5e5e5" }}>
                                        {/* Header Row */}
                                        <div
                                            style={{
                                                display: "grid",
                                                gridTemplateColumns: "1fr 1fr",
                                                gap: "16px",
                                                padding: "12px 16px",
                                                background: "#f8f9fa",
                                                borderBottom: "1px solid #e5e5e5",
                                                fontWeight: 600,
                                                fontSize: "13px",
                                                color: "#666",
                                            }}
                                        >
                                            <div>Column Name</div>
                                            <div>Column Type</div>
                                        </div>

                                        {/* Data Rows */}
                                        {existingColumns.map((col: ExistingColumn, index: number) => (
                                            <div
                                                key={col.ID || index}
                                                style={{
                                                    display: "grid",
                                                    gridTemplateColumns: "1fr 1fr",
                                                    gap: "16px",
                                                    padding: "10px 16px",
                                                    background: "#fff",
                                                    borderBottom: index < existingColumns.length - 1 ? "1px solid #f0f0f0" : "none",
                                                    fontSize: "14px",
                                                    color: "#333",
                                                }}
                                            >
                                                {/* // Aman 25/3/26 commented */}
                                                {/* <div>{col.ColumnName || col.Title || "-"}</div>
                                                <div style={{ color: "#666" }}>{col.ColumnType || "-"}</div> */}
                                                 {/* Aman 25/3/26 - Priority to IsRename, then fallback to ColumnName/Title */}
                                                    <div>
                                                        {col.IsRename || col.ColumnName || (col as any).columnname || col.Title || "No Metadata Available"}
                                                    </div>
                                                    <div style={{ color: "#666" }}>
                                                        {col.ColumnType || (col as any).columntype || "No Metadata Available"}
                                                    </div>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            )}

                            {/* NEW COLUMNS TABLE */}
                            <div>
                                <div
                                    style={{
                                        fontSize: "14px",
                                        fontWeight: 600,
                                        color: "#666",
                                        marginBottom: "8px",
                                        padding: "0 4px",
                                    }}
                                >
                                    Add New Columns
                                </div>

                                {newColumns.length > 0 && (
                                    <div style={{ background: "#fff", borderRadius: "4px", overflow: "hidden", border: "1px solid #e5e5e5" }}>
                                        {/* Header Row */}
                                        <div
                                            style={{
                                                display: "grid",
                                                gridTemplateColumns: "1fr 1fr 50px",
                                                gap: "12px",
                                                padding: "12px 16px",
                                                background: "#f8f9fa",
                                                borderBottom: "1px solid #e5e5e5",
                                                fontWeight: 600,
                                                fontSize: "13px",
                                                color: "#666",
                                            }}
                                        >
                                            <div>Field Name</div>
                                            <div>Field Type</div>
                                            <div style={{ textAlign: "center" }}>Remove</div>
                                        </div>

                                        {/* Data Rows */}
                                        {newColumns.map((col: ColumnData, index: number) => (
                                            <div
                                                key={index}
                                                style={{
                                                    display: "grid",
                                                    gridTemplateColumns: "1fr 1fr 50px",
                                                    gap: "12px",
                                                    padding: "10px 16px",
                                                    background: "#fff",
                                                    borderBottom: index < newColumns.length - 1 ? "1px solid #f0f0f0" : "none",
                                                    alignItems: "center",
                                                }}
                                            >
                                                <div>
                                                    <input
                                                        type="text"
                                                        placeholder="Enter field name"
                                                        value={col.ColumnName}
                                                        onChange={(e: React.ChangeEvent<HTMLInputElement>) =>
                                                            handleNewColumnChange(index, "ColumnName", e.target.value)
                                                        }
                                                        style={{
                                                            width: "100%",
                                                            padding: "8px 12px",
                                                            border: "1px solid #d1d1d1",
                                                            borderRadius: "4px",
                                                            fontSize: "14px",
                                                            outline: "none",
                                                            boxSizing: "border-box",
                                                            background: "#f9f9f9",
                                                        }}
                                                    />
                                                </div>

                                                <div>
                                                    <select
                                                        value={col.ColumnType}
                                                        onChange={(e: React.ChangeEvent<HTMLSelectElement>) =>
                                                            handleNewColumnChange(index, "ColumnType", e.target.value)
                                                        }
                                                        style={{
                                                            width: "100%",
                                                            padding: "8px 12px",
                                                            border: "1px solid #d1d1d1",
                                                            borderRadius: "4px",
                                                            fontSize: "14px",
                                                            outline: "none",
                                                            boxSizing: "border-box",
                                                            background: "#f9f9f9",
                                                            cursor: "pointer",
                                                        }}
                                                    >
                                                        <option>Single Line of Text</option>
                                                        <option>Multiple Line of Text</option>
                                                        <option>Number</option>
                                                        <option>Date & Time</option>
                                                        <option>Yes or No</option>
                                                    </select>
                                                </div>

                                                {/* Remove Button - Always Visible */}
                                                <div style={{ textAlign: "center" }}>
                                                    <button
                                                        type="button"
                                                        onClick={(e) => handleRemoveColumn(index, e)}
                                                        style={{
                                                            background: "#dc3545",
                                                            border: "none",
                                                            color: "#fff",
                                                            fontSize: "18px",
                                                            cursor: "pointer",
                                                            width: "32px",
                                                            height: "32px",
                                                            display: "flex",
                                                            alignItems: "center",
                                                            justifyContent: "center",
                                                            padding: 0,
                                                            fontWeight: "bold",
                                                            borderRadius: "4px",
                                                            margin: "0 auto",
                                                            transition: "background 0.2s",
                                                        }}
                                                        title="Remove this field"
                                                        onMouseEnter={(e) => {
                                                            e.currentTarget.style.background = "#c82333";
                                                        }}
                                                        onMouseLeave={(e) => {
                                                            e.currentTarget.style.background = "#dc3545";
                                                        }}
                                                    >
                                                        ×
                                                    </button>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                )}

                                {/* Empty State */}
                                {newColumns.length === 0 && (
                                    <div
                                        style={{
                                            background: "#fff",
                                            padding: "30px 20px",
                                            textAlign: "center",
                                            borderRadius: "4px",
                                            border: "1px solid #e5e5e5",
                                            color: "#999",
                                            fontSize: "14px",
                                        }}
                                    >
                                        Click the blue <strong>+</strong> button below to add new fields
                                    </div>
                                )}
                            </div>
                        </>
                    )}
                </div>

                {/* FOOTER */}
                <div
                    style={{
                        padding: "16px 20px",
                        borderTop: "1px solid #e5e5e5",
                        display: "flex",
                        justifyContent: "flex-end",
                        alignItems: "center",
                        gap: "12px",
                        background: "#fff",
                    }}
                >
                    {/* Save Button - Green Check */}
                    <button
                        type="button"
                        onClick={handleSubmit}
                        disabled={loading || newColumns.length === 0}
                        style={{
                            background: newColumns.length === 0 ? "#ccc" : "#28a745",
                            color: "#fff",
                            border: "none",
                            borderRadius: "50%",
                            width: "48px",
                            height: "48px",
                            cursor: loading || newColumns.length === 0 ? "not-allowed" : "pointer",
                            fontSize: "24px",
                            fontWeight: "bold",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            boxShadow: newColumns.length === 0 ? "none" : "0 2px 8px rgba(40, 167, 69, 0.3)",
                            opacity: loading || newColumns.length === 0 ? 0.6 : 1,
                        }}
                        title="Save Fields"
                    >
                        ✓
                    </button>

                    {/* Add Button - Blue Plus */}
                    <button
                        type="button"
                        onClick={handleAddNewColumn}
                        disabled={loading}
                        style={{
                            background: "#007bff",
                            color: "#fff",
                            border: "none",
                            borderRadius: "50%",
                            width: "48px",
                            height: "48px",
                            cursor: loading ? "not-allowed" : "pointer",
                            fontSize: "28px",
                            fontWeight: "bold",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            boxShadow: "0 2px 8px rgba(0, 123, 255, 0.3)",
                            opacity: loading ? 0.6 : 1,
                        }}
                        title="Add Field"
                    >
                        +
                    </button>
                </div>
            </div>
        </Modal.Body>
    </Modal>
);

};

export default MetadataModal;
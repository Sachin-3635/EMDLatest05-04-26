import React, { useEffect, useMemo, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory, useLocation } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { sp } from "@pnp/sp";
import logo from "../../assets/sona-comstarlogo.png";

/** ========== UI Helpers ========== */
const Section = ({ title, children }: any) => (
    <div className="form-section">
        <h3>{title}</h3>
        {children}
    </div>
);
const Grid = ({ children, style }: any) => (
    <div className="form-grid" style={style}>
        {children}
    </div>
);
const Field = ({ label, children, full }: any) => (
    <div className={full ? "form-field full" : "form-field"}>
        <label>{label}</label>
        {children}
    </div>
);

const ViewForm = (props: ISonaEmdProps) => {
    const history = useHistory();
    const location = useLocation();
    const spCrudOps = SPCRUDOPS();

    /** ---------- Setup PnP ---------- */
    useEffect(() => {
        try {
            sp.setup({ spfxContext: (props as any).context });
        } catch {
            // no-op
        }
    }, [props]);

    /** ---------- Parse ItemId (search + hash; HashRouter-safe) ---------- */
    const itemId = useMemo(() => {
        let idStr = new URLSearchParams(location.search || window.location.search).get("ItemId");
        if (!idStr) {
            const hash = window.location.hash || "";
            const qIndex = hash.indexOf("?");
            const search = qIndex >= 0 ? hash.substring(qIndex) : "";
            idStr = new URLSearchParams(search).get("ItemId");
        }
        const parsed = idStr ? parseInt(idStr, 10) : NaN;
        console.log("[Closure] Parsed ItemId:", idStr, "=>", parsed);
        return parsed;
    }, [location.search, location.hash]);

    /** ---------- State ---------- */
    const [title, setTitle] = useState<string>("");
    const [currentStatus, setCurrentStatus] = useState<string>("");

    // Requestor
    const [employee, setEmployee] = useState({
        EmployeeCode: "",
        EmployeeName: "",
        Division: "",
        Department: "",
        Location: "",
        ReportingManager: "",
        RM: "",
        HOD: "",
        ContactNo: "",
        EmployeeStatus: "",
        EmployeeEmail: "",
        ReportingManagerId: 0,
        HODId: 0,
    });

    // EMD Request Details (read-only)
    const [vendorCode, setVendorCode] = useState<string>("");
    const [vendorSite, setVendorSite] = useState<string>("");
    const [tenderNo, setTenderNo] = useState<string>("");
    const [tenderDate, setTenderDate] = useState<string>("");
    const [tenderAmount, setTenderAmount] = useState<string>("");
    const [emdAmount, setEmdAmount] = useState<string>("");
    const [emdPercentage, setEmdPercentage] = useState<string>("");
    const [vendorNameTitle, setVendorNameTitle] = useState<string>("");
    const [tenderTypeText, setTenderTypeText] = useState<string>("");
    const [currencyText, setCurrencyText] = useState<string>("");
    const [productType, setProductType] = useState<string>("");
    const [modeofPayment, setModeofPayment] = useState<string>("");
    const [tenderClosingDate, setTenderClosingDate] = useState<string>("");

    // Vouching details (AP Team) – read-only display
    const [vouchingDate, setVouchingDate] = useState<string>("");
    const [glCode, setGLCode] = useState<string>("");
    const [voucherNo, setVoucherNo] = useState<string>("");

    // Histories
    const [approverComment, setApproverComment] = useState<string>("");
    const [approverActionDate, setApproverActionDate] = useState<string>(""); // Modified fallback
    const [apTeamComment, setAPTeamComment] = useState<string>("");
    const [apActionDate, setAPActionDate] = useState<string>(""); // VouchingDate/Modified fallback
    const [treasuryComment, setTreasuryComment] = useState<string>("");

    // UTR (Treasury) – read-only for closure
    const [utrNo, setUTRNo] = useState<string>("");
    const [utrDate, setUTRDate] = useState<string>("");

    // Attachments (existing)
    const [attachments, setAttachments] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);

    // Stage-wise grouped attachments for Supporting Docs
    const [apDocs, setApDocs] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);
    const [treasuryDocs, setTreasuryDocs] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);
    const [manacDocs, setManacDocs] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);

    // ---------- Closure Action (required) ----------
    const [dateOfReceipt, setDateOfReceipt] = useState<string>("");
    const [bankAccount, setBankAccount] = useState<string>("");
    const [closureAmount, setClosureAmount] = useState<string>("");
    const [closureComments, setClosureComments] = useState<string>(""); // maps to ARComment
    const [files, setFiles] = useState<File[]>([]); // new attachments
    const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
    // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&
    const [isTenderDuplicate, setIsTenderDuplicate] = useState<boolean>(false);
    // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&
    // Prevent duplicate submit clicks
    const [isSaving, setIsSaving] = useState<boolean>(false);
    const [wfHistory, setWfHistory] = useState<any[]>([]);
    /** ---------- Helpers ---------- */
    const formatINR = (val?: string) => {
        if (!val) return "";
        const n = Number((val || "").toString().replace(/,/g, ""));
        if (isNaN(n)) return val || "";
        return n.toLocaleString("en-IN");
    };

    function onFilesPicked(e: React.ChangeEvent<HTMLInputElement>) {
        const fl = Array.from(e.target.files || []);
        setFiles(fl);
    }

    function groupAttachmentsByStage(list: Array<{ FileName: string; ServerRelativeUrl: string }>) {
        const toStage = (f: string) => {
            const name = (f || "").toLowerCase();
            if (name.startsWith("treasury_")) return "treasury";
            if (name.startsWith("ap_") || name.startsWith("apteam_")) return "ap";
            if (name.startsWith("manac_") || name.startsWith("manak_")) return "manac";
            return "other";
        };
        const ap: typeof list = [];
        const treasury: typeof list = [];
        const manac: typeof list = [];
        list.forEach((x) => {
            const stage = toStage(x.FileName);
            if (stage === "ap") ap.push(x);
            else if (stage === "treasury") treasury.push(x);
            else if (stage === "manac") manac.push(x);
        });
        setApDocs(ap);
        setTreasuryDocs(treasury);
        setManacDocs(manac);
    }

    async function loadAttachments(id: number) {
        try {
            const files = await sp.web.lists.getByTitle("EMDDetails").items.getById(id).attachmentFiles();
            console.log("[Closure] Attachments:", files);
            setAttachments(files || []);
            groupAttachmentsByStage(files || []);
        } catch (e) {
            console.warn("[Closure] Attachments fetch failed", e);
            setAttachments([]);
            setApDocs([]);
            setTreasuryDocs([]);
            setManacDocs([]);
        }
    }


    const formatDate = (date: any) => {
        if (!date) return "-";

        // Handle DD/MM/YYYY
        const parts = date.split("/");
        if (parts.length === 3) {
            const [day, month, year] = parts;
            const d = new Date(`${year}-${month}-${day}`);

            return isNaN(d.getTime())
                ? "-"
                : d.toLocaleDateString("en-GB");
        }

        return "-";
    };

    // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&
    const checkDuplicateTendorNumber = async (tenderNo: string) => {
        try {
            const spx = await spCrudOps;

            const items = await spx.getData(
                "EMDDetails",
                "Id,TenderNo",
                "",
                // `TenderNo eq '${tenderNo}'`,
                `TenderNo eq '${tenderNo}' and Status ne 'Rejected'`,
                { column: "Id", isAscending: false },
                1,
                props
            );

            // 🔥 Important: current item ko exclude karo
            // if (items.length > 0 && items[0].Id !== itemId) {
            //     return true;
            // }
            return items.length > 1;

            // return false;
        } catch (error) {
            console.error("Duplicate check error:", error);
            return false;
        }
    };
    useEffect(() => {
        const checkDuplicate = async () => {
            if (tenderNo) {
                const isDuplicate = await checkDuplicateTendorNumber(tenderNo);
                setIsTenderDuplicate(isDuplicate);
            }
        };

        checkDuplicate();
    }, [tenderNo]);

    // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&

    /** ---------- Robust fetch ---------- */
    async function fetchEMDDetailsRobust(id: number) {
        const spx = await spCrudOps;

        const select =
            "ID,Title,Status,Modified," +
            "EmployeeName,EmployeeCode,Department,Division,Location,RM,HOD,ContactNo,EmployeeStatus,Email," +
            "VendorCode,VendorSite," +
            "ProductType/ProductType" +
            "ModeofPayment/Mode" +
            "TenderNo,TenderDate,TenderAmount,EMDAmount,EMDPercentage,TenderClosingDate," +
            "TenderTypeId,TenderType/TenderType," +
            "VendorNameId,VendorName/Title,VendorName/Name," +
            "CurrencyId,Currency/Currency," +
            "VouchingDate,GLCode,VoucherNo,APTeamComment," +
            "ApproverComment," +
            "UTRNo,UTRDate,TreasuryComment," +
            "DateofReceipt,BankAccount,Amount,ARComment,ApprovalMatrix,WFHistory";

        const expand = "VendorName,TenderType,Currency";

        const tries = [
            { filter: `ID eq ${id}`, order: { column: "ID", isAscending: true } },
            { filter: `Id eq ${id}`, order: { column: "Id", isAscending: true } },
        ];

        for (const t of tries) {
            try {
                console.log("[Closure] getData with", t.filter);
                const res = await spx.getData("EMDDetails", select, expand, t.filter, t.order, 1, props);
                if (res && res.length) return res[0];
            } catch (e) {
                console.warn("[Closure] getData error with", t.filter, e);
            }
        }

        try {
            console.log("[Closure] Fallback PnP getById");
            const pnp = await sp.web.lists
                .getByTitle("EMDDetails")
                .items.getById(id)
                .select(
                    "ID",
                    "Title",
                    "Status",
                    "Modified",
                    "EmployeeName",
                    "EmployeeCode",
                    "Department",
                    "Division",
                    "Location",
                    "RM",
                    "HOD",
                    "ContactNo",
                    "EmployeeStatus",
                    "Email",
                    "VendorCode",
                    "VendorSite",
                    "TenderNo",
                    "TenderDate",
                    "TenderAmount",
                    "EMDAmount",
                    "ModeofPayment/Id",
                    "ModeofPayment/Mode",
                    "ProductType/Id",
                    "ProductType/ProductType",
                    "EMDPercentage",
                    "TenderClosingDate",
                    "TenderTypeId",
                    "VendorNameId",
                    "CurrencyId",
                    "VouchingDate",
                    "GLCode",
                    "VoucherNo",
                    "APTeamComment",
                    "ApproverComment",
                    "UTRNo",
                    "UTRDate",
                    "TreasuryComment",
                    "DateofReceipt",
                    "BankAccount",
                    "Amount",
                    "ARComment",
                    "Currency/Currency",
                    "VendorName/Name",
                    "VendorName/Title",
                    "TenderType/TenderType",
                    "ApprovalMatrix",
                    "WFHistory"
                )
                .expand("VendorName", "TenderType", "Currency", "ModeofPayment", "ProductType")
                .get();
            return pnp;
        } catch (e) {
            console.error("[Closure] PnP fallback error", e);
            return null as any;
        }
    }

    /** ---------- Load details ---------- */
    useEffect(() => {
        if (!itemId || Number.isNaN(itemId)) {
            console.warn("[Closure] Invalid ItemId", itemId);
            return;
        }

        void (async () => {
            const details = await fetchEMDDetailsRobust(itemId);
            if (!details) {
                alert("No data found for this ItemId.");
                return;
            }

            console.log("[Closure] Details:", details);

            // Summary
            setTitle(details.Title || "");
            setCurrentStatus(details.Status || "");

            // Requestor
            setEmployee({
                EmployeeCode: details.EmployeeCode || "",
                EmployeeName: details.EmployeeName || "",
                Division: details.Division || "",
                Department: details.Department || "",
                Location: details.Location || "",
                ReportingManager: details.ReportingManager || "",
                HOD: details.HOD || "",
                RM: details.RM || "",
                ContactNo: details.ContactNo || "",
                EmployeeStatus: details.EmployeeStatus || "",
                EmployeeEmail: details.Email || "",
                ReportingManagerId: details.ReportingManagerId || 0,
                HODId: details.HODId || 0,
            });

            // EMD details
            setVendorCode(details.VendorCode || "");
            setVendorSite(details.VendorSite || "");
            setTenderNo(details.TenderNo || "");
            setTenderDate(details.TenderDate ? String(details.TenderDate).split("T")[0] : "");
            setTenderAmount(details.TenderAmount || "");
            setEmdAmount(details.EMDAmount || "");
            setTenderClosingDate(details.TenderClosingDate ? String(details.TenderClosingDate).split("T")[0] : "");
            if (details.TenderAmount && details.EMDAmount) {
                const t = parseFloat(details.TenderAmount || "0");
                const e = parseFloat(details.EMDAmount || "0");
                setEmdPercentage(t > 0 ? ((e / t) * 100).toFixed(2) : "");
            } else {
                setEmdPercentage(details.EMDPercentage || "");
            }

            // Lookups
            setVendorNameTitle(details.VendorName?.Name || "");
            setTenderTypeText(details.TenderType?.TenderType || "");
            setCurrencyText(details.Currency?.Currency || "");
            setProductType(details.ProductType?.ProductType || "");
            setModeofPayment(details.ModeofPayment?.Mode);
            // Vouching details (AP)
            setVouchingDate(details.VouchingDate ? String(details.VouchingDate).split("T")[0] : "");
            setGLCode(details.GLCode || "");
            setVoucherNo(details.VoucherNo || "");
            setAPTeamComment(details.APTeamComment || "");
            const apD = details.VouchingDate || details.Modified;
            setAPActionDate(apD ? String(apD).split("T")[0] : "");

            // Approver history
            setApproverComment(details.ApproverComment || "");
            const appD = details.Modified;
            setApproverActionDate(appD ? String(appD).split("T")[0] : "");

            // Treasury (UTR)
            setUTRNo(details.UTRNo || "");
            setUTRDate(details.UTRDate ? String(details.UTRDate).split("T")[0] : "");
            setTreasuryComment(details.TreasuryComment || "");

            // Prefill closure action if already present
            setDateOfReceipt(details.DateofReceipt ? String(details.DateofReceipt).split("T")[0] : "");
            setBankAccount(details.BankAccount != null ? String(details.BankAccount) : "");
            setClosureAmount(details.Amount != null ? String(details.Amount) : "");
            setClosureComments(details.ARComment || "");

            // 🔥 LOAD APPROVAL MATRIX
            try {
                let matrix: any[] = [];

                if (details.ApprovalMatrix) {
                    matrix =
                        typeof details.ApprovalMatrix === "string"
                            ? JSON.parse(details.ApprovalMatrix)
                            : details.ApprovalMatrix;
                }

                // ✅ sort by sequence
                matrix.sort((a: any, b: any) => a.Seq - b.Seq);

                console.log("🔥 Approval Matrix:", matrix); // debug

                setApprovalMatrix(matrix);
                // 🔥 LOAD WF HISTORY
                try {
                    let history: any[] = [];

                    if (details.WFHistory) {
                        history =
                            typeof details.WFHistory === "string"
                                ? JSON.parse(details.WFHistory)
                                : details.WFHistory;
                    }

                    setWfHistory(history || []);

                    console.log("🔥 WFHistory:", history);

                } catch (e) {
                    console.error("WFHistory parse error", e);
                    setWfHistory([]);
                }

            } catch (e) {
                console.error("Error parsing ApprovalMatrix", e);
                setApprovalMatrix([]);
            }

            // Attachments
            await loadAttachments(itemId);
        })();
    }, [itemId]);

    /** ---------- Validation Helpers ---------- */
    const isNumeric = (x: string) => /^-?\d+(\.\d{1,2})?$/.test((x || "").toString().trim());
    const canClose = () => {
        if (!dateOfReceipt) return false;
        if (!bankAccount.trim() || !isNumeric(bankAccount)) return false;
        if (!closureAmount.trim() || !isNumeric(closureAmount)) return false;
        if (!closureComments.trim()) return false;
        return true;
    };

    /** ---------- Submit Closure ---------- */
    // const onCloseSubmit = async () => {
    //     if (isSaving) return;

    //     try {
    //         setIsSaving(true);

    //         if (!itemId || Number.isNaN(itemId)) return alert("Invalid item id.");

    //         // Final status after AR Closure submission (EXACT choice text)
    //         const NEXT_STATUS = "Pending for Closure Approval";

    //         // Optional safety: allow close only from EMD Paid
    //         // if (currentStatus !== "EMD Paid") {
    //         //   const proceed = confirm(
    //         //     `Current Status is "${currentStatus}". Close is typically after "EMD Paid". Proceed?`
    //         //   );
    //         //   if (!proceed) { setIsSaving(false); return; }
    //         // }

    //         if (!dateOfReceipt) return alert("Date of Receipt is required.");
    //         if (!bankAccount.trim()) return alert("Bank Account is required.");
    //         if (!isNumeric(bankAccount)) return alert("Bank Account must be a number.");
    //         if (!closureAmount.trim()) return alert("Amount is required.");
    //         if (!isNumeric(closureAmount)) return alert("Amount must be a valid number (upto 2 decimals).");
    //         if (!closureComments.trim()) return alert("Comments are required.");

    //         const spx = await spCrudOps;

    //         // Update EMDDetails with closure information
    //         await spx.updateData(
    //             "EMDDetails",
    //             itemId,
    //             {
    //                 DateofReceipt: dateOfReceipt,     // Date
    //                 BankAccount: Number(bankAccount), // Number
    //                 Amount: Number(closureAmount),    // Number
    //                 ARComment: closureComments,       // Multiline
    //                 Status: NEXT_STATUS,              // Choice
    //             },
    //             props
    //         );

    //         // Upload new attachments if enabled
    //         // if (files.length > 0) {
    //         //   await Promise.all(
    //         //     files.map((file) => spx.addAttchmentInList(file, "EMDDetails", itemId, file.name, props))
    //         //   );
    //         // }

    //         alert("✅ Submitted for Closure Approval.");
    //         // Redirect to Closure Approval dashboard
    //         history.push("/ClosureApprovalARDashboard");
    //     } catch (e) {
    //         console.error("[Closure] Submit error", e);
    //         alert("Something went wrong. Please check console.");
    //     } finally {
    //         setIsSaving(false);
    //     }
    // };

    const onCloseSubmit = async () => {
        if (isSaving) return;

        try {
            setIsSaving(true);

            if (!itemId || Number.isNaN(itemId)) return alert("Invalid item id.");

            if (!dateOfReceipt) return alert("Date of Receipt is required.");
            if (!bankAccount.trim()) return alert("Bank Account is required.");
            if (!closureAmount.trim()) return alert("Amount is required.");
            if (!closureComments.trim()) return alert("Comments are required.");

            const spx = await spCrudOps;

            // 🔥 1️⃣ GET CURRENT ITEM
            const items = await spx.getData(
                "EMDDetails",
                "Id,ApprovalMatrix,WFHistory",
                "",
                `Id eq ${itemId}`,
                { column: "Id", isAscending: false },
                1,
                props
            );

            if (!items || items.length === 0) {
                alert("Item not found");
                return;
            }

            const item = items[0];

            // 🔥 2️⃣ PARSE OLD MATRIX
            let oldMatrix: any[] = [];

            try {
                oldMatrix =
                    typeof item.ApprovalMatrix === "string"
                        ? JSON.parse(item.ApprovalMatrix)
                        : item.ApprovalMatrix || [];
            } catch {
                oldMatrix = [];
            }

            // 🔥 3️⃣ GET MASTER MATRIX (EMDApprovalMatrix)
            const master = await spx.getData(
                "EMDApprovalMatrix",
                "ID,Role/RoleName,Approver/ID,Approver/Title",
                "Role,Approver",
                "",
                { column: "ID", isAscending: true },
                1000,
                props
            );

            // 🔥 4️⃣ BUILD UPDATED MATRIX
            let matrix = master.map((x: any, index: number) => {

                const oldItem = oldMatrix.find(
                    (o) => o.Role === x.Role?.RoleName
                );

                const newApproverID = x.Approver?.ID;

                // ✅ IF APPROVER CHANGED → RESET
                if (!oldItem || String(oldItem.ApproverID) !== String(newApproverID)) {
                    return {
                        Seq: index + 1,
                        Role: x.Role?.RoleName,
                        ApproverID: newApproverID,
                        Approver: x.Approver?.Title,
                        Status: "Not Started"
                    };
                }

                // ✅ KEEP OLD STATUS
                return {
                    ...oldItem,
                    ApproverID: newApproverID,
                    Approver: x.Approver?.Title
                };
            });

            // 🔥 5️⃣ KEEP RM/HOD FROM OLD MATRIX
            oldMatrix.forEach(oldItem => {
                if (!matrix.find((x: { Role: any; }) => x.Role === oldItem.Role)) {
                    matrix.push(oldItem);
                }
            });

            // 🔥 6️⃣ SORT
            matrix.sort((a: { Seq: number; }, b: { Seq: number; }) => a.Seq - b.Seq);

            // 🔥 7️⃣ SET FIRST PENDING
            let foundPending = false;

            matrix = matrix.map((item: { Status: string; }) => {
                if (!foundPending && item.Status !== "Approved") {
                    foundPending = true;
                    return { ...item, Status: "Pending" };
                }

                if (item.Status !== "Approved") {
                    return { ...item, Status: "Not Started" };
                }

                return item;
            });

            // 🔥 8️⃣ CURRENT APPROVER
            const current = matrix.find((x: { Status: string; }) => x.Status === "Pending");

            // ===============================
            // 🔥 WF HISTORY UPDATE
            // ===============================
            let prevHistory: any[] = [];

            try {
                prevHistory =
                    typeof item.WFHistory === "string"
                        ? JSON.parse(item.WFHistory)
                        : item.WFHistory || [];
            } catch {
                prevHistory = [];
            }

            // ✅ Current user
            const currentUser =
                props.context?.pageContext?.user?.displayName || "AR Team";

            // ✅ New entry
            const newEntry = {
                CurrentApprover: currentUser,
                ActionTaken: "Closure Submitted",
                Comment: closureComments,
                Date: new Date().toISOString()
            };

            // ✅ PUSH
            prevHistory.push(newEntry);

            // ✅ SAVE
            const wfHistoryPayload = JSON.stringify(prevHistory);

            console.log("🔥 FINAL WFHistory:", prevHistory);

            // 🔥 9️⃣ UPDATE SHAREPOINT
            await spx.updateData(
                "EMDDetails",
                itemId,
                {
                    DateofReceipt: dateOfReceipt,
                    BankAccount: Number(bankAccount),
                    Amount: Number(closureAmount),
                    ARComment: closureComments,
                    Status: "Pending for Closure Approval",

                    // ✅ IMPORTANT
                    ApprovalMatrix: JSON.stringify(matrix),
                    //  CurrentApproverId: current?.ApproverID || null,
                    PendingAt: current ? `Pending at ${current.Approver}` : "Completed",
                    WFHistory: wfHistoryPayload

                },
                props
            );

            alert("✅ Submitted for Closure Approval");

            history.push("/ClosureApprovalARDashboard");

        } catch (e) {
            console.error("[Closure] Submit error", e);
            alert("Something went wrong.");
        } finally {
            setIsSaving(false);
        }
    };
    return (
        <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
            <div className='row'>
                <div className='col-md-12'>
                    <div className='Main-Boxpoup'>
                        {/* 🔹 Header */}
                        <div className="bordered">
                            <img src={logo} />
                            <h1>View Details</h1>
                            {title ? <div style={{ marginTop: 4, color: "#666" }}>Ref: {title}</div> : null}
                        </div>
                        <div className="approvalFlow">

                            {/* Initiator */}
                            <ul >
                                <li className="flowStep green">
                                    Initiator
                                </li>
                            </ul>

                            {approvalMatrix.map((step, index) => {

                                let stepClass = "grey";

                                const firstPending = approvalMatrix.findIndex(s => s.Status === "Pending");
                                const isLastStep = index === approvalMatrix.length - 1;

                                // 🔴 Rejected
                                if (step.Status === "Rejected") {
                                    stepClass = "red";
                                }

                                // 🟢 Approved
                                else if (step.Status === "Approved") {
                                    stepClass = "green";
                                }

                                // 🟢 Last Accepted
                                else if (isLastStep && step.Status === "Accepted") {
                                    stepClass = "green";
                                }

                                // 🟠 Current Pending
                                else if (index === firstPending) {
                                    stepClass = "orange";
                                }

                                return (

                                    <ul key={index}>
                                        <li className={`flowStep ${stepClass}`}>
                                            {step.ApproverName || step.ApproverName || step.Approver || step.Role}
                                        </li>
                                    </ul>
                                );
                            })}

                        </div>
                        <div className='borderedbox'>
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Requestor Information</label>
                            </div>
                            <div className='main-formcontainer'>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Code" className='font'>Employee Code</label> : &nbsp;&nbsp;
                                        <label className='fonttext'> {employee.EmployeeCode} </label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {employee.EmployeeName}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Email" className='font'>Employee Email </label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {employee.EmployeeEmail}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Contact No" className='font'>Contact No</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {employee.ContactNo}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Employee Status" className='font'>Employee Status</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {employee.EmployeeStatus}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="Division" className='font'>Division</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {employee.Division}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {employee.Location}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {employee.ReportingManager}</label>
                                    </div>
                                    <div className='col-md-4'>
                                        <label htmlFor="HOD" className='font'>HOD</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {employee.HOD}</label>
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label htmlFor="Location" className='font'>Department</label> : &nbsp;&nbsp;
                                        <label className='fonttext'>  {employee.Department}</label>
                                    </div>
                                </div>
                            </div>
                            {isTenderDuplicate && (
                                <section>
                                    <h5 style={{ color: "green" }}>
                                        This Tender No. is available with another EMD request.
                                    </h5>
                                </section>
                            )}
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>EMD Request Details</label>
                            </div>
                            <div className='main-formcontainer'>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label className='font'>Vendor Code </label>
                                        <input value={vendorCode} className='form-control readonly readonly' readOnly />
                                    </div>
                                    <div className='col-md-4'>
                                        <label className='font'>Vendor Name </label>
                                        <input value={vendorNameTitle} className='form-control readonly readonly' readOnly />

                                    </div>
                                    <div className='col-md-4'>
                                        <label className='font'>Vendor Site </label>.
                                        <input value={vendorSite} readOnly className="form-control readonly readonly" />
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    {/* <div className='col-md-4'>
                                    <label className="font">Contract Type </label>
                                    <input value={tenderNo} readOnly className="form-control readonly readonly" />
                                </div> */}
                                    <div className='col-md-4'>
                                        <label className="font">Tender No </label>
                                        <input value={tenderNo} readOnly className="form-control readonly readonly" />

                                    </div>
                                    <div className='col-md-4'>
                                        <label className="font">Tender Date </label>
                                        <input type="date" value={tenderDate} readOnly className="form-control readonly readonly" />
                                    </div>
                                    <div className='col-md-4'>
                                        <label className="font">Tender Type </label>
                                        <input value={tenderTypeText} readOnly className="form-control readonly readonly" />
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label className="font">Tender Amount </label>
                                        <input value={formatINR(tenderAmount)} readOnly className="form-control readonly readonly" />
                                    </div>
                                    <div className='col-md-4'>
                                        <label className="font">EMD Amount </label>
                                        <input value={formatINR(emdAmount)} readOnly className="form-control readonly readonly" />
                                    </div>
                                    <div className='col-md-4'>
                                        <label className="font">Currency </label>
                                        <input value={currencyText} readOnly className="form-control readonly readonly" />
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label className="font">Tender Closing Date </label>
                                        <input type="date" value={tenderClosingDate} readOnly className="form-control readonly readonly" />
                                    </div>
                                    <div className='col-md-4'>
                                        <label className="font">EMD Percentage </label>
                                        <input value={emdPercentage} className="form-control readonly readonly" readOnly />
                                    </div>
                                    <div className='col-md-4'>
                                        <label className="font">Mode of Payment </label>
                                        <input value={productType} readOnly className="form-control readonly readonly" />
                                    </div>
                                </div>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label className="font">Product Type </label>
                                        <input type="text" value={modeofPayment} readOnly className="form-control readonly readonly" />
                                    </div>
                                </div>
                            </div>
                            {vouchingDate && (
                                <>
                                    <div className="heading1" style={{ marginTop: "10px" }}>
                                        <label>Vouching Details</label>
                                    </div>
                                    <div className='main-formcontainer'>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Vouching Date </label>
                                                <input type="date" className='form-control readonly readonly' value={vouchingDate} readOnly />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>GL </label>
                                                <input value={glCode} readOnly className='form-control readonly readonly' />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>vendor Code </label>
                                                <input value={vendorCode} readOnly className="form-control readonly readonly" />
                                            </div>
                                        </div>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className="font">voucher No. </label>
                                                <input value={voucherNo} readOnly className="form-control readonly" />
                                            </div>
                                        </div>
                                    </div>
                                </>
                            )}
                            {utrNo && (
                                <>
                                    <div className="heading1" style={{ marginTop: "10px" }}>
                                        <label>UTR Details</label>
                                    </div>
                                    <div className='main-formcontainer'>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>UTR No </label>
                                                <input className='form-control readonly' value={utrNo} readOnly />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>UTR Date </label>
                                                <input type="date" value={utrDate} readOnly className='form-control readonly' />
                                            </div>
                                        </div>
                                    </div>
                                </>
                            )}
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Work Flow History</label>
                            </div>
                            <div className='main-formcontainer'>
                                <div className='row mb-20'>
                                    <div className='col-md-12'>
                                        <div className="overflow-x-auto">
                                            <table className="custom-table">
                                                <thead>
                                                    <tr>
                                                        <th className="px-4 py-2">Action By</th>
                                                        <th className="px-4 py-2">Action Taken</th>
                                                        <th className="px-4 py-2">Date</th>
                                                        <th className="px-4 py-2">Comment</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {wfHistory.length > 0 ? (
                                                        wfHistory.map((item, index) => {

                                                            // const formatDate = (date: any) => {
                                                            //   if (!date) return "-";
                                                            //   const d = new Date(date);
                                                            //   return isNaN(d.getTime())
                                                            //     ? "-"
                                                            //     : d.toLocaleString("en-GB");
                                                            // };

                                                            const formatDate = (date: any) => {
                                                                if (!date) return "-";

                                                                // Handle DD/MM/YYYY
                                                                const parts = date.split("/");
                                                                if (parts.length === 3) {
                                                                    const [day, month, year] = parts;
                                                                    const d = new Date(`${year}-${month}-${day}`);

                                                                    return isNaN(d.getTime())
                                                                        ? "-"
                                                                        : d.toLocaleDateString("en-GB");
                                                                }

                                                                return "-";
                                                            };

                                                            return (
                                                                <tr key={index}>
                                                                    <td className="px-4 py-2">{item.CurrentApprover}</td>
                                                                    <td className="px-4 py-2">{item.ActionTaken}</td>
                                                                    <td className="px-4 py-2">{formatDate(item.Date)}</td>
                                                                    <td className="px-4 py-2">{item.Comment || "-"}</td>
                                                                </tr>
                                                            );
                                                        })
                                                    ) : (
                                                        <tr>
                                                            <td colSpan={4}>No history available</td>
                                                        </tr>
                                                    )}
                                                </tbody>

                                            </table>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div className="heading1" style={{ marginTop: "10px" }}>
                                <label>Upload Documents</label>
                            </div>
                            <div className='main-formcontainer'>
                                <div className='row mb-20'>
                                    <div className='col-md-4'>
                                        <label className='font'> Existing Attachments </label>
                                        {attachments.length === 0 ? (
                                            <div>-</div>
                                        ) : (
                                            <ul style={{ margin: 0, paddingLeft: 18 }}>
                                                {attachments.map((a) => (
                                                    <li key={a.ServerRelativeUrl}>
                                                        <a href={a.ServerRelativeUrl} target="_blank" rel="noreferrer">
                                                            {a.FileName}
                                                        </a>
                                                    </li>
                                                ))}
                                            </ul>
                                        )}
                                    </div>
                                </div>
                            </div>
                            {dateOfReceipt && (
                                <>
                                    <div className="heading1" style={{ marginTop: "10px" }}>
                                        <label>EMD Updated</label>
                                    </div>
                                    <div className='main-formcontainer'>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Date of Receipt <span className="Mantorystar">*</span> </label>
                                                <input type="date" className='form-control readonly' value={dateOfReceipt} readOnly />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>Bank Account <span className="Mantorystar">*</span> </label>
                                                <input value={bankAccount} readOnly className='form-control readonly' />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>Amount <span className="Mantorystar">*</span> </label>
                                                <input value={closureAmount} readOnly className='form-control readonly' />
                                            </div>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>UTR No <span className="Mantorystar">*</span> </label>
                                                    <input value={utrNo} readOnly className='form-control readonly' />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </>
                            )}
                            <div className='row my-3'>
                                <div className='col-md-12'>
                                    <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                                        <button className="reset-btn" onClick={() => history.goBack()}>Exit</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default ViewForm;
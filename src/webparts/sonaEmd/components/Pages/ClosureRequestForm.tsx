import React, { useEffect, useMemo, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory, useLocation } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { sp } from "@pnp/sp";

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

const ClosureRequestForm = (props: ISonaEmdProps) => {
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

    // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&
    const [isTenderDuplicate, setIsTenderDuplicate] = useState<boolean>(false);
    // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&

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
            "TenderNo,TenderDate,TenderAmount,EMDAmount,EMDPercentage,TenderClosingDate," +
            "TenderTypeId,TenderType/TenderType," +
            "VendorNameId,VendorName/Title," +
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
                    "VendorName/Title",
                    "TenderType/TenderType",
                    "ApprovalMatrix",
                    "WFHistory"
                )
                .expand("VendorName", "TenderType", "Currency")
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

            try {
                let matrix: any[] = [];

                // =========================
                // 🔹 LOAD OLD MATRIX
                // =========================
                try {
                    const parsed =
                        typeof details.ApprovalMatrix === "string"
                            ? JSON.parse(details.ApprovalMatrix)
                            : details.ApprovalMatrix;

                    matrix = Array.isArray(parsed) ? parsed : [];
                } catch {
                    matrix = [];
                }

                console.log("🔥 Old Matrix:", matrix);

                const spx = await spCrudOps;

                // =========================
                // 🔹 FETCH CLOSURE APPROVERS
                // =========================
                const closureApprovers = await spx.getData(
                    "EMDApprovalMatrix",
                    "ID,Role/RoleName,Approver/ID,Approver/Title",
                    "Role,Approver",
                    "RequestType eq 'EMD Closure Approval'",
                    { column: "ID", isAscending: true },
                    1000,
                    props
                );

                console.log("🔥 Closure Approvers:", closureApprovers);

                // 🚨 DEBUG CHECK
                if (!closureApprovers || closureApprovers.length === 0) {
                    console.warn("⚠ No Closure Approvers Found");
                }

                // =========================
                // 🔹 ADD NEW APPROVERS (NO DUPLICATE ISSUE)
                // =========================
                let maxSeq =
                    matrix.length > 0
                        ? Math.max(...matrix.map((x: any) => x.Seq || 0))
                        : 0;

                closureApprovers.forEach((x: any, index: number) => {

                    // 🔥 USE ApproverID instead of Role (FIX)
                    const exists = matrix.some(
                        (m: any) => m.ApproverID === x.Approver?.ID
                    );

                    if (!exists) {
                        maxSeq += 1; // 🔥 IMPORTANT FIX

                        matrix.push({
                            Seq: maxSeq,
                            Role: x.Role?.RoleName,
                            Approver: x.Approver?.Title,
                            ApproverID: x.Approver?.ID,
                            Status: "Not Started"
                        });
                    }
                });

                // =========================
                // 🔹 NORMALIZE
                // =========================
                matrix = matrix.map((x: any, index: number) => ({
                    Seq: x.Seq ?? index + 1,
                    Role: x.Role || "",
                    Approver: x.Approver || x.ApproverName || "",
                    ApproverID: x.ApproverID || x.Id || null,
                    Status: x.Status || "Not Started"
                }));

                // =========================
                // 🔹 SORT
                // =========================
                matrix.sort((a: any, b: any) => (a.Seq || 0) - (b.Seq || 0));

                console.log("🔥 FINAL MATRIX:", matrix);

                setApprovalMatrix(matrix);

            } catch (e) {
                console.error("Matrix load error", e);
                setApprovalMatrix([]);
            }
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

        if (!dateOfReceipt) { alert("Date of Receipt is required."); return; }
        if (!bankAccount.trim()) { alert("Bank Account is required."); return; }
        if (!closureAmount.trim()) { alert("Amount is required."); return; }
        if (!closureComments.trim()) { alert("Comments are required."); return; }

        const spx = await spCrudOps;

        // =========================
        // 🔥 1️⃣ GET CURRENT ITEM
        // =========================
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

        // =========================
        // 🔥 2️⃣ PARSE OLD MATRIX
        // =========================
        let oldMatrix: any[] = [];

        try {
            oldMatrix =
                typeof item.ApprovalMatrix === "string"
                    ? JSON.parse(item.ApprovalMatrix)
                    : item.ApprovalMatrix || [];
        } catch {
            oldMatrix = [];
        }

        console.log("🔥 Old Matrix:", oldMatrix);

        // =========================
        // 🔥 3️⃣ GET CLOSURE APPROVERS
        // =========================
        const closureApprovers = await spx.getData(
            "EMDApprovalMatrix",
            "ID,Role/RoleName,Approver/ID,Approver/Title",
            "Role,Approver",
            "RequestType eq 'EMD Closure Approval'",
            { column: "ID", isAscending: true },
            1000,
            props
        );

        console.log("🔥 Closure Approvers:", closureApprovers);

        // =========================
        // 🔥 4️⃣ BUILD MATRIX
        // =========================
        let matrix = [...oldMatrix];

        let maxSeq =
            matrix.length > 0
                ? Math.max(...matrix.map((x: any) => x.Seq || 0))
                : 0;

        closureApprovers.forEach((x: any) => {

            // ✅ FIX: check by ApproverID (NOT Role)
            const exists = matrix.some(
                (m: any) => m.ApproverID === x.Approver?.ID
            );

            if (!exists) {
                maxSeq += 1;

                matrix.push({
                    Seq: maxSeq,
                    Role: x.Role?.RoleName,
                    ApproverID: x.Approver?.ID,
                    Approver: x.Approver?.Title,
                    Status: "Not Started"
                });
            }
        });

        // =========================
        // 🔥 5️⃣ RESET REJECTED / SENT BACK
        // =========================
        matrix = matrix.map((item: any) => {
            if (item.Status === "Rejected" || item.Status === "Sent Back") {
                return { ...item, Status: "Not Started" };
            }
            return item;
        });

        // =========================
        // 🔥 6️⃣ REORDER RM & HOD FIRST
        // =========================
        const normalizeRole = (r: any) =>
            String(r || "").trim().toLowerCase();

        const rmItem = matrix.find((x: any) => normalizeRole(x.Role) === "rm");
        const hodItem = matrix.find((x: any) => normalizeRole(x.Role) === "hod");

        const remaining = matrix.filter((x: any) =>
            normalizeRole(x.Role) !== "rm" && normalizeRole(x.Role) !== "hod"
        );

        matrix = [
            ...(rmItem ? [rmItem] : []),
            ...(hodItem ? [hodItem] : []),
            ...remaining
        ].map((item: any, index: number) => ({
            ...item,
            Seq: index + 1
        }));

        // =========================
        // 🔥 7️⃣ FORCE FIRST PENDING (FIXED)
        // =========================
        let foundPending = false;

        matrix = matrix.map((item: any) => {
            if (!foundPending && item.Status !== "Approved") {
                foundPending = true;
                return { ...item, Status: "Pending" };
            }
            if (item.Status !== "Approved") {
                return { ...item, Status: "Not Started" };
            }
            return item;
        });

        // =========================
        // 🔥 8️⃣ CURRENT APPROVER
        // =========================
        const finalCurrent =
            matrix.find((x: any) => x.Status === "Pending") ||
            matrix.find((x: any) => x.Status === "Not Started") ||
            null;

        console.log("👉 Current Approver:", finalCurrent);

        // =========================
        // 🔥 9️⃣ WORKFLOW HISTORY
        // =========================
        let prevHistory: any[] = [];

        try {
            prevHistory =
                typeof item.WFHistory === "string"
                    ? JSON.parse(item.WFHistory)
                    : item.WFHistory || [];
        } catch {
            prevHistory = [];
        }

        const currentUser =
            props.context?.pageContext?.user?.displayName || "AR Team";

        const today = new Date();
        const formattedDate =
            String(today.getDate()).padStart(2, '0') + '/' +
            String(today.getMonth() + 1).padStart(2, '0') + '/' +
            today.getFullYear();

        prevHistory.push({
            CurrentApprover: currentUser,
            ActionTaken: "Closure Submitted",
            Comment: closureComments,
            Date: formattedDate
        });

        // =========================
        // 🔥 🔟 UPDATE SHAREPOINT
        // =========================
        await spx.updateData(
            "EMDDetails",
            itemId,
            {
                DateofReceipt: dateOfReceipt,
                BankAccount: Number(bankAccount),
                Amount: Number(closureAmount),
                ARComment: closureComments,

                Status: "Pending for Closure Approval",

                ApprovalMatrix: JSON.stringify(matrix),
                CurrentApproverId: finalCurrent?.ApproverID || null,
                PendingAt: finalCurrent
                    ? `Pending at ${finalCurrent.Approver}`
                    : "Completed",

                WFHistory: JSON.stringify(prevHistory)
            },
            props
        );

        alert("✅ Submitted for Closure Approval");
        history.goBack();

    } catch (e) {
        console.error("[Closure] Submit error", e);
        alert("Something went wrong.");
    } finally {
        setIsSaving(false);
    }
};
    return (
        <div className="forex-wrapper">
            {/* ================= HEADER ================= */}
            <div className="forex-header">
                <h2>Closure Request Form </h2>
                {title ? <div style={{ marginTop: 4, color: "#666" }}>Ref: {title}</div> : null}
            </div>

            <div className="forex-card">

                {/* ================= APPROVAL HIERARCHY ================= */}
                {/* <div className="emd-hierarchy">
                    <div className="emd-step active-step">{employee.EmployeeName}</div>

                    <div className="emd-step" style={{ marginLeft: "30px" }}>{employee.ReportingManager}</div>

                    <div className="emd-step" style={{ marginLeft: "30px" }}>{employee.HOD}</div>
                </div> */}
                <div className="headerApproval">
                    <div className="approvalFlow">

                        {/* Initiator */}
                        <div className="flowStep green">Initiator</div>

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
                                <div key={index} className={`flowStep ${stepClass}`}>
                                    {/* 🔥 SHOW NAME INSTEAD OF ROLE */}
                                    {/* {step.Approver || step.Role} */}
                                    {step.ApproverName || step.ApproverName || step.Approver || step.Role}
                                </div>
                            );
                        })}

                    </div>
                </div>

                {/* Requestor Information */}
                <Section title="Requestor Information">
                    <Grid style={{ marginTop: "20px" }}>
                        <Field label="Employee Code">
                            <input value={employee.EmployeeCode} readOnly />
                        </Field>
                        <Field label="Employee Name">
                            <input value={employee.EmployeeName} readOnly />
                        </Field>
                        <Field label="Division">
                            <input value={employee.Division} readOnly />
                        </Field>
                        <Field label="Location">
                            <input value={employee.Location} readOnly />
                        </Field>
                        <Field label="Reporting Manager">
                            <input value={employee.RM || employee.ReportingManager} readOnly />
                        </Field>
                        <Field label="HOD">
                            <input value={employee.HOD} readOnly />
                        </Field>
                        <Field label="Contact No">
                            <input value={employee.ContactNo} readOnly />
                        </Field>
                        <Field label="Employee Status">
                            <input value={employee.EmployeeStatus} readOnly />
                        </Field>
                        <Field label="Department">
                            <input value={employee.Department} readOnly />
                        </Field>
                        <Field label="Employee Email" full>
                            <input type="email" value={employee.EmployeeEmail} readOnly />
                        </Field>
                    </Grid>
                </Section>

                {isTenderDuplicate && (
                    <section>
                        <h5 style={{ color: "green" }}>
                            This Tender No. is available with another EMD request.
                        </h5>
                    </section>
                )}

                {/* EMD Request Details */}
                <Section title="EMD Request Details">
                    <Grid>
                        <Field label="Vendor Code">
                            <input value={vendorCode} readOnly />
                        </Field>
                        <Field label="Vendor Name">
                            <input value={vendorNameTitle} readOnly />
                        </Field>
                        <Field label="Vendor Site">
                            <input value={vendorSite} readOnly />
                        </Field>

                        <Field label="Tender No.">
                            <input value={tenderNo} readOnly />
                        </Field>
                        <Field label="Tender Date">
                            <input type="date" value={tenderDate} readOnly />
                        </Field>
                        <Field label="Tender Type">
                            <input value={tenderTypeText} readOnly />
                        </Field>

                        <Field label="Tender Amount">
                            <input value={formatINR(tenderAmount)} readOnly />
                        </Field>
                        <Field label="EMD Amount">
                            <input value={formatINR(emdAmount)} readOnly />
                        </Field>
                        <Field label="Currency">
                            <input value={currencyText} readOnly />
                        </Field>

                        <Field label="Tender Closing Date">
                            <input type="date" value={tenderClosingDate} readOnly />
                        </Field>
                        <Field label="EMD Percentage">
                            <input value={emdPercentage} readOnly />
                        </Field>
                    </Grid>
                </Section>

                {/* Vouching Details (AP) */}
                <Section title="Vouching Details">
                    <Grid>
                        <Field label="Vouching Date*">
                            <input type="date" value={vouchingDate} readOnly />
                        </Field>
                        <Field label="GL*">
                            <input value={glCode} readOnly />
                        </Field>
                        <Field label="Vendor Code*">
                            <input value={vendorCode} readOnly />
                        </Field>
                        <Field label="Voucher No.*">
                            <input value={voucherNo} readOnly />
                        </Field>
                    </Grid>
                </Section>

                {/* UTR Details */}
                <Section title="UTR Details">
                    <Grid>
                        <Field label="UTR No*">
                            <input value={utrNo} readOnly />
                        </Field>
                        <Field label="UTR Date*">
                            <input type="date" value={utrDate} readOnly />
                        </Field>
                    </Grid>
                </Section>

                <Section title="Workflow History">
                    <div className="wfTableWrapper">
                        <table className="wfTable">
                            <thead>
                                <tr>
                                    <th>Action By</th>
                                    <th>Action</th>
                                    <th>Date</th>
                                    <th>Comment</th>
                                </tr>
                            </thead>
                            <tbody>
                                {/* {wfHistory.length > 0 ? (
                                    wfHistory.map((item, i) => (                                       

                                        <tr key={i}>
                                            <td>{item.CurrentApprover}</td>
                                            <td>{item.ActionTaken}</td>
                                            <td>{new Date(item.Date).toLocaleString("en-GB")}</td>
                                            <td>{item.Comment || "-"}</td>
                                        </tr>
                                    ))
                                ) : (
                                    <tr>
                                        <td colSpan={4}>No history available</td>
                                    </tr>
                                )} */}
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
                                                <td>{item.CurrentApprover}</td>
                                                <td>{item.ActionTaken}</td>
                                                <td>{formatDate(item.Date)}</td>
                                                <td>{item.Comment || "-"}</td>
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
                </Section>
                {/* Workflow History */}
                {/* <Section title="Workflow History">
                    <Grid>
                        <Field label="UTR Details entered by">
                            <input value="Treasury Team" readOnly />
                        </Field>
                        <Field label="Action Taken">
                            <input value="UTR details" readOnly />
                        </Field>
                        <Field label="Action Date">
                            <input value={utrDate || "-"} readOnly />
                        </Field>

                        <Field label="Vouching Details entered by">
                            <input value="AP Team" readOnly />
                        </Field>
                        <Field label="Action Taken">
                            <input value="Vouching details" readOnly />
                        </Field>
                        <Field label="Action Date">
                            <input value={apActionDate || "-"} readOnly />
                        </Field>

                        <Field label="Approval by">
                            <input value="MANAC Team" readOnly />
                        </Field>
                        <Field label="Action Taken">
                            <input value="Approved" readOnly />
                        </Field>
                        <Field label="Action Date">
                            <input value={approverActionDate || "-"} readOnly />
                        </Field>
                    </Grid>
                </Section> */}

                {/* Comment History */}
                {/* <Section title="Comment History">
                    <Grid>
                        <Field label="Comment By">
                            <input value="Treasury Team" readOnly />
                        </Field>
                        <Field label="Comment">
                            <textarea value={treasuryComment || "UTR details updated"} readOnly rows={3} />
                        </Field>
                        <Field label="Comment Date">
                            <input value={utrDate || "-"} readOnly />
                        </Field>

                        <Field label="Comment By">
                            <input value="AP Team" readOnly />
                        </Field>
                        <Field label="Comment">
                            <textarea value={apTeamComment || "Vouching details updated"} readOnly rows={3} />
                        </Field>
                        <Field label="Comment Date">
                            <input value={apActionDate || "-"} readOnly />
                        </Field>

                        <Field label="Comment By">
                            <input value="MANAC Team" readOnly />
                        </Field>
                        <Field label="Comment">
                            <textarea value={approverComment || "Approving for Vouching"} readOnly rows={3} />
                        </Field>
                        <Field label="Comment Date">
                            <input value={approverActionDate || "-"} readOnly />
                        </Field>
                    </Grid>
                </Section> */}

                {/* Uploaded Documents (all) */}
                <Section title="Uploaded Documents">
                    <Grid>
                        <Field label="Existing Attachments">
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
                        </Field>
                    </Grid>
                </Section>

                {/* Action (Closure) */}
                <Section title="Action">
                    <Grid>
                        <Field
                            //  label="Date of Receipt*"
                            label={<>
                                Date of Receipt.<span style={{ color: "red" }}>*</span>
                            </>
                            }

                        >
                            <input type="date" value={dateOfReceipt} onChange={(e) => setDateOfReceipt(e.target.value)} />
                        </Field>
                        <Field
                            //  label="Bank Account*"
                            label={<>
                                Bank Account.<span style={{ color: "red" }}>*</span>
                            </>
                            }
                        >
                            <input
                                type="number"
                                step="1"
                                value={bankAccount}
                                onChange={(e) => setBankAccount(e.target.value)}
                                placeholder="Enter bank account"
                            />
                        </Field>
                        <Field
                            // label="Amount*"
                            label={<>
                                Amount.<span style={{ color: "red" }}>*</span>
                            </>
                            }

                        >
                            <input
                                type="number"
                                step="0.01"
                                value={closureAmount}
                                onChange={(e) => setClosureAmount(e.target.value)}
                                placeholder="Enter amount"
                            />
                        </Field>
                        <Field label="UTR No">
                            <input value={utrNo} readOnly />
                        </Field>
                        <Field
                            // label="Comments*" 
                            label={<>
                                Comments.<span style={{ color: "red" }}>*</span>
                            </>
                            }

                            full>
                            <textarea
                                rows={3}
                                value={closureComments}
                                onChange={(e) => setClosureComments(e.target.value)}
                                placeholder="Enter comments"
                            />
                        </Field>

                        {/* <Field label="Attach" full>
              <input type="file" multiple onChange={onFilesPicked} />
              {files.length > 0 && (
                <div style={{ marginTop: 6 }}>
                  {files.map((f) => (
                    <span key={f.name} style={{ marginRight: 8, fontSize: 12 }}>
                      {f.name}
                    </span>
                  ))}
                </div>
              )}
              <div style={{ marginTop: 8, color: "#6a6a6a", fontSize: 12 }}>
                Tip: Stage-wise Supporting Docs prefix: <code>Treasury_*</code>, <code>AP_*</code>, <code>MANAC_*</code>.
              </div>
            </Field> */}
                    </Grid>

                    {/* Buttons */}
                    <div className="button-row">
                        {/* <button className="btn-submit" disabled={isSaving || !canClose()} onClick={onCloseSubmit}>
                            {isSaving ? "Saving..." : "Close"}
                        </button> */}
                        <button className="btn-submit" onClick={onCloseSubmit}>
                            {isSaving ? "Saving..." : "Close"}
                        </button>
                        <button className="btn-exit" onClick={() => history.push("/TreasuryLandingPage")}>
                            Exit
                        </button>
                    </div>
                </Section>
            </div>
        </div>
    );
};

export default ClosureRequestForm;
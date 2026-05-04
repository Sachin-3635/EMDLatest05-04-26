import React, { useEffect, useMemo, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory, useLocation } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { sp } from "@pnp/sp";

/** ===========================
 *   ROUTES & STATUS (CONFIG)
 *  =========================== */
// AP team dashboard (after Approve)
const ROUTE_AP_CLOSURE_DASHBOARD = "/APTeamDashboardClosedByAR";

// Initiator edit form route
const ROUTE_INITIATOR_EDIT_BASE = "/NewRequest"; // <-- yaha apna initiator edit form ka path set karein
const INITIATOR_EDIT_USES_QUERY = true; // true => /NewRequest?ItemId=7&mode=edit ; false => /NewRequest/7?mode=edit

// Reject ko bhi 'Send back' jaisa chalayen (as per your ask)
const REJECT_BEHAVES_AS_SEND_BACK = true;

// Exact SharePoint choice texts
const STATUS = {
  PENDING_AR_APPROVAL: "Pending for Closure Approval", // item comes in this state
  APPROVED_NEXT: "Pending for Closure Vouching",       // Approve par ye set hoga
  SEND_BACK: "Send back",
  REJECTED: "Rejected",
} as const;

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

const ARClosureApprovalForm = (props: ISonaEmdProps) => {
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
    console.log("[AR-Closure] Parsed ItemId:", idStr, "=>", parsed);
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

  // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&
  const [isTenderDuplicate, setIsTenderDuplicate] = useState<boolean>(false);
  // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&

  // Extra lookups visible in screenshot
  const [modeOfPaymentText, setModeOfPaymentText] = useState<string>("");
  const [contractTypeText, setContractTypeText] = useState<string>("");
  const [productTypeText, setProductTypeText] = useState<string>("");

  // Vouching details (AP Team) – read-only display
  const [vouchingDate, setVouchingDate] = useState<string>("");
  const [glCode, setGLCode] = useState<string>("");
  const [voucherNo, setVoucherNo] = useState<string>("");

  // Histories
  const [approverComment, setApproverComment] = useState<string>("");
  const [approverActionDate, setApproverActionDate] = useState<string>("");
  const [apTeamComment, setAPTeamComment] = useState<string>("");
  const [apActionDate, setAPActionDate] = useState<string>("");
  const [treasuryComment, setTreasuryComment] = useState<string>("");

  // UTR (Treasury)
  const [utrNo, setUTRNo] = useState<string>("");
  const [utrDate, setUTRDate] = useState<string>("");

  // Attachments
  const [attachments, setAttachments] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);

  // Closure (read-only here)
  const [dateOfReceipt, setDateOfReceipt] = useState<string>("");
  const [bankAccount, setBankAccount] = useState<string>("");
  const [closureAmount, setClosureAmount] = useState<string>("");
  const [closureComments, setClosureComments] = useState<string>(""); // ARComment

  // Approver action
  const [approvalRemark, setApprovalRemark] = useState<string>("");
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [wfHistory, setWfHistory] = useState<any[]>([]);

  /** ---------- Helpers ---------- */
  const formatINR = (val?: string) => {
    if (!val) return "";
    const n = Number((val || "").toString().replace(/,/g, ""));
    if (isNaN(n)) return val || "";
    return n.toLocaleString("en-IN");
  };

  async function loadAttachments(id: number) {
    try {
      const files = await sp.web.lists.getByTitle("EMDDetails").items.getById(id).attachmentFiles();
      setAttachments(files || []);
    } catch (e) {
      console.warn("[AR-Closure] Attachments fetch failed", e);
      setAttachments([]);
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
        `TenderNo eq '${tenderNo}' and Status ne 'Rejected' and Id ne ${itemId}`,
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

  /** ---------- Robust fetch (extended lookups) ---------- */
  async function fetchEMDDetailsRobust(id: number) {
    const spx = await spCrudOps;

    const select =
      "ID,Title,Status,Modified," +
      // Requestor
      "EmployeeName,EmployeeCode,Division,Department,Location,RM,HOD,ContactNo,EmployeeStatus,Email," +
      // EMD details
      "VendorCode,VendorSite," +
      "TenderNo,TenderDate,TenderAmount,EMDAmount,EMDPercentage,TenderClosingDate," +
      "TenderTypeId,TenderType/TenderType,TenderType/Title," +
      "VendorNameId,VendorName/Title,VendorName/Name," +
      "CurrencyId,Currency/Currency,Currency/Title," +
      "ModeofPaymentId,ModeofPayment/Title," +
      "ContractTypeId,ContractType/Title," +
      "ProductTypeId,ProductType/Title," +
      // Vouching (AP)
      "VouchingDate,GLCode,VoucherNo,APTeamComment," +
      // Approver (MANAC)
      "ApproverComment," +
      // Treasury (UTR)
      "UTRNo,UTRDate,TreasuryComment," +
      // Closure (AR)
      "DateofReceipt,BankAccount,Amount,ARComment,ApprovalMatrix,WFHistory";

    const expand = "VendorName,TenderType,Currency,ModeofPayment,ContractType,ProductType";

    const tries = [
      { filter: `ID eq ${id}`, order: { column: "ID", isAscending: true } },
      { filter: `Id eq ${id}`, order: { column: "Id", isAscending: true } },
    ];

    for (const t of tries) {
      try {
        const res = await spx.getData("EMDDetails", select, expand, t.filter, t.order, 1, props);
        if (res && res.length) return res[0];
      } catch (e) {
        console.warn("[AR-Closure] getData error with", t.filter, e);
      }
    }

    // Fallback PnP
    try {
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
          "ModeofPaymentId",
          "ContractTypeId",
          "ProductTypeId",
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
          "Currency/Title",
          "VendorName/Id",
          "VendorName/Name",
          "TenderType/TenderType",
          "TenderType/Title",
          "ModeofPayment/Title",
          "ContractType/Title",
          "ProductType/Title"
        )
        .expand("VendorName", "TenderType", "Currency", "ModeofPayment", "ContractType", "ProductType")
        .get();
      return pnp;
    } catch (e) {
      console.error("[AR-Closure] PnP fallback error", e);
      return null as any;
    }
  }

  /** ---------- Load details ---------- */
  useEffect(() => {
    if (!itemId || Number.isNaN(itemId)) {
      console.warn("[AR-Closure] Invalid ItemId", itemId);
      return;
    }

    void (async () => {
      const details = await fetchEMDDetailsRobust(itemId);
      if (!details) {
        alert("No data found for this ItemId.");
        return;
      }

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
      setTenderTypeText(details.TenderType?.Title || details.TenderType?.TenderType || "");
      setCurrencyText(details.Currency?.Title || details.Currency?.Currency || "");
      setModeOfPaymentText(details.ModeofPayment?.Title || "");
      setContractTypeText(details.ContractType?.Title || "");
      setProductTypeText(details.ProductType?.Title || "");

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

      // Closure (AR)
      setDateOfReceipt(details.DateofReceipt ? String(details.DateofReceipt).split("T")[0] : "");
      setBankAccount(details.BankAccount != null ? String(details.BankAccount) : "");
      setClosureAmount(details.Amount != null ? String(details.Amount) : "");
      setClosureComments(details.ARComment || "");
      // 🔥 APPROVAL MATRIX LOAD
      try {
        let matrix: any[] = [];

        if (details.ApprovalMatrix) {
          matrix =
            typeof details.ApprovalMatrix === "string"
              ? JSON.parse(details.ApprovalMatrix)
              : details.ApprovalMatrix;
        }

        matrix.sort((a: any, b: any) => a.Seq - b.Seq);

        console.log("🔥 ApprovalMatrix:", matrix);

        setApprovalMatrix(matrix);


      } catch (e) {
        console.error("ApprovalMatrix parse error", e);
        setApprovalMatrix([]);
      }
      // Attachments
      await loadAttachments(itemId);
      // ===============================
      // 🔥 LOAD WF HISTORY
      // ===============================
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
    })();
  }, [itemId]);

  /** ---------- Navigation helpers ---------- */
  const navigateToInitiatorEdit = (id: number) => {
    if (INITIATOR_EDIT_USES_QUERY) {
      history.push({ pathname: ROUTE_INITIATOR_EDIT_BASE, search: `?ItemId=${id}&mode=edit` });
    } else {
      history.push(`${ROUTE_INITIATOR_EDIT_BASE}/${id}?mode=edit`);
    }
  };

  /** ---------- Update helpers ---------- */
  const requireRemark = (s: string) => s === STATUS.SEND_BACK || s === STATUS.REJECTED;

  const getUpdatedApprovalMatrix = async (oldMatrix: any[]) => {

    try {
      const sp = await spCrudOps;

      // 🔥 GET LATEST APPROVER MASTER
      const master = await sp.getData(
        "EMDApprovalMatrix",
        "ID,Role/RoleName,Approver/ID,Approver/Title",
        "Role,Approver",
        "",
        { column: "ID", isAscending: true },
        1000,
        props
      );

      if (!master || master.length === 0) {
        console.warn("⚠ No master approval matrix found");
        return oldMatrix;
      }

      // 🔥 CREATE NEW MATRIX FROM MASTER
      let newMatrix = master.map((m: any, index: number) => ({
        Seq: index + 1,
        Role: m.Role?.RoleName,
        Approver: m.Approver?.Title,
        ApproverID: String(m.Approver?.ID),
        Status: "Not Started"
      }));

      // 🔥 MERGE WITH OLD MATRIX
      newMatrix = newMatrix.map((newItem: { Role: any; Seq: any; }) => {

        const oldItem = oldMatrix.find(
          (x) => x.Role === newItem.Role || x.Seq === newItem.Seq
        );

        // ✅ RM / HOD PRESERVE (important for you)
        if (oldItem && (oldItem.Role === "RM" || oldItem.Role === "HOD")) {
          return oldItem;
        }

        return {
          ...newItem,
          Status: oldItem?.Status || "Not Started"
        };
      });

      // 🔥 SORT
      newMatrix.sort((a: any, b: any) => a.Seq - b.Seq);

      return newMatrix;

    } catch (e) {
      console.error("❌ Error updating approval matrix", e);
      return oldMatrix;
    }
  };

  // const handleAction = async (action: "APPROVE" | "SEND_BACK" | "REJECT") => {
  //   if (isSaving) return;
  //   if (!itemId || Number.isNaN(itemId)) return alert("Invalid item id.");

  //   try {
  //     // Optional: enforce current state
  //     // if (currentStatus !== STATUS.PENDING_AR_APPROVAL) {
  //     //   return alert(`This request is not in "${STATUS.PENDING_AR_APPROVAL}" state.`);
  //     // }

  //     let newStatus = "";
  //     let afterUpdateNav: () => void = () => { };

  //     if (action === "APPROVE") {
  //       // ✅ Approve -> Pending for Closure Vouching -> AP Team dashboard
  //       newStatus = STATUS.APPROVED_NEXT;
  //       afterUpdateNav = () => history.push(ROUTE_AP_CLOSURE_DASHBOARD);
  //     }

  //     if (action === "SEND_BACK") {
  //       // ✅ Send Back -> Initiator edit
  //       newStatus = STATUS.SEND_BACK;
  //       if (!approvalRemark.trim()) return alert("Comment is required for Send Back.");
  //       afterUpdateNav = () => navigateToInitiatorEdit(itemId);
  //     }

  //     if (action === "REJECT") {
  //       // ✅ As per your ask: Reject bhi Initiator ko wapas jaaye 'Send back' ki tarah
  //       if (REJECT_BEHAVES_AS_SEND_BACK) {
  //         newStatus = STATUS.SEND_BACK;
  //       } else {
  //         newStatus = STATUS.REJECTED;
  //       }
  //       if (!approvalRemark.trim()) return alert("Comment is required for Reject.");
  //       afterUpdateNav = () => navigateToInitiatorEdit(itemId);
  //     }

  //     setIsSaving(true);
  //     const spx = await spCrudOps;

  //     await spx.updateData(
  //       "EMDDetails",
  //       itemId,
  //       {
  //         Status: newStatus,
  //         ApproverComment: approvalRemark || approverComment || "",
  //       },
  //       props
  //     );

  //     alert(`✅ Status updated to "${newStatus}".`);
  //     afterUpdateNav();
  //   } catch (e) {
  //     console.error("[AR-Closure] Action error", e);
  //     alert("Something went wrong. Please check console.");
  //   } finally {
  //     setIsSaving(false);
  //   }
  // };
  const handleAction = async (action: "APPROVE" | "SEND_BACK" | "REJECT") => {

    if (isSaving) return;
    if (!itemId || Number.isNaN(itemId)) return alert("Invalid item id.");

    try {

      if ((action === "SEND_BACK" || action === "REJECT") && !approvalRemark.trim()) {
        return alert("Comment is required.");
      }

      setIsSaving(true);
      const spx = await spCrudOps;

      // 🔥 1. GET ITEM WITH MATRIX
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

      // 🔥 2. PARSE MATRIX
      let matrix: any[] = [];

      try {
        matrix =
          typeof item.ApprovalMatrix === "string"
            ? JSON.parse(item.ApprovalMatrix)
            : item.ApprovalMatrix || [];
      } catch {
        matrix = [];
      }

      if (!Array.isArray(matrix) || matrix.length === 0) {
        alert("Approval matrix empty");
        return;
      }

      // New when got error of no approver found
      // const oldPending = matrix.find(x => x.Status === "Pending");
      // 🔥 3. UPDATE FROM MASTER (IMPORTANT)
      // new when got no approver found
      // matrix = await getUpdatedApprovalMatrix(matrix);
      // 🔥 restore pending
      // if (oldPending) {
      //   let matchIndex = matrix.findIndex(
      //     x => x.Role?.trim().toLowerCase() === oldPending.Role?.trim().toLowerCase()
      //   );

      //   // 🔥 fallback (important)
      //   if (matchIndex === -1) {
      //     matchIndex = matrix.findIndex(
      //       x => Number(x.ApproverID) === Number(oldPending.ApproverID)
      //     );
      //   }

      //   if (matchIndex !== -1) {
      //     matrix = matrix.map((m, i) => ({
      //       ...m,
      //       Status: i === matchIndex
      //         ? "Pending"
      //         : m.Status === "Pending"
      //           ? "Not Started"
      //           : m.Status
      //     }));
      //   }
      // }
      // 🔥 SORT
      matrix.sort((a: any, b: any) => a.Seq - b.Seq);

      // 🔥 FIND CURRENT
      const currentIndex = matrix.findIndex(x => x.Status === "Pending");

      if (currentIndex === -1) {
        alert("No pending approver");
        return;
      }

      let payload: any = {};

      // ============================
      // ✅ APPROVE
      // ============================
      // ============================
      // ✅ APPROVE (UPDATED)
      // ============================
      if (action === "APPROVE") {

        // 🔥 Mark current as Approved
        matrix[currentIndex].Status = "Approved";

        // 🔥 FIND AP PERFORMER
        const apPerformerIndex = matrix.findIndex(
          (x) => x.Role?.toLowerCase().includes("performer")
        );

        if (apPerformerIndex !== -1) {

          // 🔥 RESET ALL AFTER CURRENT
          matrix = matrix.map((m, index) => {
            if (index > currentIndex) {
              return { ...m, Status: "Not Started" };
            }
            return m;
          });

          // 🔥 SET AP PERFORMER AS PENDING
          matrix[apPerformerIndex].Status = "Pending";

          const apPerformer = matrix[apPerformerIndex];

          payload.CurrentApproverId = Number(apPerformer.ApproverID);
          payload.Status = STATUS.APPROVED_NEXT;
          payload.PendingAt = `Pending at ${apPerformer.Approver}`;

        } else {

          // ❌ If AP Performer not found
          console.warn("AP Performer not found in matrix");

          payload.CurrentApproverId = null;
          payload.Status = "Completed";
          payload.PendingAt = "Completed";
        }
      }

      // ============================
      // 🔴 SEND BACK / REJECT
      // ============================
      // else {

      //   matrix[currentIndex].Status = "Rejected";

      //   //payload.CurrentApproverId = null;
      //   payload.Status = "EMD Paid";
      //   payload.PendingAt = `Rejected by ${matrix[currentIndex].Approver}`;
      // }

      else if (action === "SEND_BACK") {

        matrix[currentIndex].Status = "Sent Back";

        payload.CurrentApproverId = null;
        payload.Status = "Sent Back";
        payload.PendingAt = `Sent back by ${matrix[currentIndex].Approver}`;

      }
      else if (action === "REJECT") {

        matrix[currentIndex].Status = "Rejected";

        matrix = matrix.map((m, i) => ({
          ...m,
          Status: i < currentIndex ? m.Status : (i === currentIndex ? "Rejected" : "Not Started")
        }));

        payload.CurrentApproverId = null;
        // payload.Status = "Rejected";
        payload.Status = "EMD Paid";
        payload.PendingAt = `Rejected by ${matrix[currentIndex].Approver}`;

        // 🔥 ADD THIS (clear fields)
        payload.DateofReceipt = null;
        payload.BankAccount = null;
        payload.Amount = null;
        payload.ARComment = "";

      }

      // 🔥 SAVE MATRIX
      payload.ApprovalMatrix = JSON.stringify(matrix);

      // 🔥 COMMENT
      payload.ApproverComment = approvalRemark || approverComment || "";

      // 🔥 DEBUG (VERY IMPORTANT)
      console.log("🔥 FINAL MATRIX:", matrix);
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

      const currentUser =
        props.context?.pageContext?.user?.displayName || "AR Approver";

      const today = new Date();

      const formattedDate =
        String(today.getDate()).padStart(2, '0') + '/' +
        String(today.getMonth() + 1).padStart(2, '0') + '/' +
        today.getFullYear();

      // const actionTaken =
      //   action === "APPROVE"
      //     ? "Approved"
      //     : "Sent Back";
      let actionTaken = "";

      if (action === "APPROVE") actionTaken = "Approved";
      else if (action === "SEND_BACK") actionTaken = "Sent Back";
      else if (action === "REJECT") actionTaken = "Rejected";

      const newEntry = {
        CurrentApprover: currentUser,
        ActionTaken: actionTaken,
        Comment: approvalRemark,
        Date: formattedDate
      };

      prevHistory.push(newEntry);

      // ✅ ADD TO PAYLOAD
      payload.WFHistory = JSON.stringify(prevHistory);

      console.log("🔥 FINAL WFHistory:", prevHistory);

      // 🔥 UPDATE SHAREPOINT
      await spx.updateData("EMDDetails", itemId, payload, props);

      alert("✅ Action completed successfully");

      // 🔥 NAVIGATION
      // if (action === "APPROVE") {
      //   history.push(ROUTE_AP_CLOSURE_DASHBOARD);
      // } else {
      //   history.push(ROUTE_AP_CLOSURE_DASHBOARD);
      // }
      history.goBack();

    } catch (e) {
      console.error("[AR-Closure] Action error", e);
      alert("Something went wrong. Please check console.");
    } finally {
      setIsSaving(false);
    }
  };
  return (
    <div className="forex-wrapper">
      {/* ================= HEADER ================= */}
      <div className="forex-header">
        <h2>Closure Form (Approval) </h2>
        {title ? <div style={{ marginTop: 4, color: "#666" }}>Ref: {title}</div> : null}
        {currentStatus ? (
          <div style={{ marginTop: 4, color: "#999", fontSize: 12 }}>
            Current Status: <b>{currentStatus}</b>
          </div>
        ) : null}
      </div>

      {/* ================= APPROVAL HIERARCHY ================= */}
      {/* <div className="emd-hierarchy">
        <div className="emd-step active-step">{employee.EmployeeName}</div>

        <div className="emd-step" style={{ marginLeft: "30px" }}>{employee.ReportingManager}</div>

        <div className="emd-step" style={{ marginLeft: "30px" }}>{employee.HOD}</div>
      </div> */}

      <div className="headerApproval">
        <div className="approvalFlow">

          {/* Initiator */}
          <div className="flowStep green">
            {employee.EmployeeName}
          </div>

          {approvalMatrix.map((step, index) => {

            let stepClass = "grey";

            const firstPending = approvalMatrix.findIndex(s => s.Status === "Pending");

            if (step.Status === "Approved") {
              stepClass = "green";
            }
            else if (step.Status === "Rejected") {
              stepClass = "red";
            }
            else if (index === firstPending) {
              stepClass = "orange"; // current step
            }

            return (
              <div key={index} className={`flowStep ${stepClass}`}>
                {/* {step.Approver || step.Role} */}
                {step.ApproverName || step.ApproverName || step.Approver || step.Role}
              </div>
            );
          })}

        </div>
      </div>
      <div className="forex-card">
        {/* Requestor Information */}
        <Section title="Requestor Information">
          <Grid style={{ marginTop: "20px" }}>
            <Field label="Employee Code">
              <input value={employee.EmployeeCode} readOnly />
            </Field>
            <Field label="Employee Name">
              <input value={employee.EmployeeName} readOnly />
            </Field>
            <Field label="Email">
              <input value={employee.EmployeeEmail} readOnly />
            </Field>
            <Field label="Contact No">
              <input value={employee.ContactNo} readOnly />
            </Field>
            <Field label="Division">
              <input value={employee.Division} readOnly />
            </Field>
            <Field label="Location">
              <input value={employee.Location} readOnly />
            </Field>
            <Field label="RM">
              <input value={employee.RM || employee.ReportingManager} readOnly />
            </Field>
            <Field label="Employee Status">
              <input value={employee.EmployeeStatus} readOnly />
            </Field>
            <Field label="Department">
              <input value={employee.Department} readOnly />
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
            <Field label="Vendor Name">
              <input value={vendorNameTitle} readOnly />
            </Field>
            <Field label="Vendor Site">
              <input value={vendorSite} readOnly />
            </Field>
            <Field label="Vendor Code">
              <input value={vendorCode} readOnly />
            </Field>

            <Field label="Tender Type">
              <input value={tenderTypeText} readOnly />
            </Field>
            <Field label="Tender No.">
              <input value={tenderNo} readOnly />
            </Field>
            <Field label="Tender Date">
              <input type="date" value={tenderDate} readOnly />
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

            <Field label="Mode of Payment">
              <input value={modeOfPaymentText} readOnly />
            </Field>
            <Field label="Contract Type">
              <input value={contractTypeText} readOnly />
            </Field>
            <Field label="Product Type">
              <input value={productTypeText} readOnly />
            </Field>

            <Field label="Tender Closing Date">
              <input type="date" value={tenderClosingDate} readOnly />
            </Field>
            <Field label="EMD Percentage">
              <input value={emdPercentage} readOnly />
            </Field>
          </Grid>
        </Section>

        {/* Uploaded Documents */}
        <Section title="Uploaded Documents">
          <Grid>
            <Field label="Existing Attachments" full>
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

        {/* Closure Details (entered by AR before approval) */}
        <Section title="Closure Details">
          <Grid>
            <Field label="Date of Receipt*">
              <input type="date" value={dateOfReceipt} readOnly />
            </Field>
            <Field label="Bank Account*">
              <input value={bankAccount} readOnly />
            </Field>
            <Field label="Amount*">
              <input value={closureAmount} readOnly />
            </Field>
            <Field label="Comments*" full>
              <textarea rows={3} value={closureComments} readOnly />
            </Field>
          </Grid>
        </Section>

        {/* Workflow History */}
        {/* <Section title="Workflow History">
          <Grid>
            <Field label="Closure Details entered by">
              <input value="Identified User" readOnly />
            </Field>
            <Field label="Action Taken">
              <input value="Closure details" readOnly />
            </Field>
            <Field label="Action Date">
              <input value={dateOfReceipt || "-"} readOnly />
            </Field>

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

            <Field label="Approval By">
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
              <input value="Identified User" readOnly />
            </Field>
            <Field label="Comment">
              <textarea value={closureComments || "Closure details updated"} readOnly rows={3} />
            </Field>
            <Field label="Comment Date">
              <input value={dateOfReceipt || "-"} readOnly />
            </Field>

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

        {/* Approver Action Panel */}
        <Section title="Approval Action">
          <Grid>
            <Field label="Comments" full>
              <textarea
                rows={3}
                value={approvalRemark}
                onChange={(e) => setApprovalRemark(e.target.value)}
                placeholder="Enter approval / reject comments"
              />
            </Field>
          </Grid>

          <div className="button-row">
            <button
              className="btn-submit"
              disabled={isSaving}
              onClick={() => handleAction("APPROVE")}
              title={`Set status to "${STATUS.APPROVED_NEXT}"`}
            >
              {isSaving ? "Saving..." : "Approve"}
            </button>

            {/* <button
              className="btn-neutral"
              disabled={isSaving}
              onClick={() => handleAction("SEND_BACK")}
              title={`Set status to "${STATUS.SEND_BACK}" (comment required)`}
            >
              Send Back
            </button> */}

            <button
              className="btn-danger"
              disabled={isSaving}
              onClick={() => handleAction("REJECT")}
              title={`Send back to initiator (comment required)`}
            >
              Reject
            </button>

            <button className="btn-exit" onClick={() => history.push("/ClosureApprovalARDashboard")}>
              Exit
            </button>
          </div>
        </Section>
      </div>
    </div>
  );
};

export default ARClosureApprovalForm;
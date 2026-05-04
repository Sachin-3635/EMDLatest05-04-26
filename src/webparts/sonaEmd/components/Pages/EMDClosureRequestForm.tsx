import React, { useEffect, useMemo, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory, useLocation } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { sp } from "@pnp/sp";

/** ===========================
 *   ROUTES & STATUS (CONFIG)
 *  =========================== */
// Close ke baad kidhar le jaana hai? (Change as needed)
const ROUTE_AFTER_CLOSE = "/APTeamDashboardClosedByAR"; // ya "/EMDClosureDashboard" / "/ClosedEMDDashboard"

// Exact SharePoint choice text (Closed)
const FINAL_STATUS = "Closed";

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

function toInputDateTimeLocal(value?: string): string {
  // Converts "2026-03-12T10:30:00Z" or "2026-03-12T10:30:00" to "2026-03-12T10:30" for <input type="datetime-local">
  if (!value) return "";
  const d = new Date(value);
  if (isNaN(d.getTime())) return "";
  const pad = (n: number) => String(n).padStart(2, "0");
  const yyyy = d.getFullYear();
  const mm = pad(d.getMonth() + 1);
  const dd = pad(d.getDate());
  const hh = pad(d.getHours());
  const mi = pad(d.getMinutes());
  return `${yyyy}-${mm}-${dd}T${hh}:${mi}`;
}

function toISOStringFromLocal(dtLocal: string): string {
  // Expects "YYYY-MM-DDTHH:mm" (local time) from <input type="datetime-local"> and returns ISO string
  // e.g., "2026-03-12T10:30" -> 2026-03-12T05:00:00.000Z (depending on timezone)
  // SharePoint DateTime columns accept ISO strings.
  const d = new Date(dtLocal);
  if (isNaN(d.getTime())) return dtLocal; // fallback as-is
  return d.toISOString();
}

const EMDClosureRequestForm = (props: ISonaEmdProps) => {
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
    console.log("[EMD-Closure] Parsed ItemId:", idStr, "=>", parsed);
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
  const [modeOfPaymentText, setModeOfPaymentText] = useState<string>("");
  const [contractTypeText, setContractTypeText] = useState<string>("");
  const [productTypeText, setProductTypeText] = useState<string>("");

  // Vouching details (display)
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

  // Closure details (entered earlier / read-only)
  const [dateOfReceipt, setDateOfReceipt] = useState<string>("");
  const [bankAccount, setBankAccount] = useState<string>("");
  const [closureAmount, setClosureAmount] = useState<string>("");
  const [closureComments, setClosureComments] = useState<string>("");
  // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&
  const [isTenderDuplicate, setIsTenderDuplicate] = useState<boolean>(false);
  // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&

  // ⬇️ New bottom action inputs (user will enter these)
  // Internal names (as given): ClosureVoucherNo (Text), ClosureVouchingDate (DateTime)
  const [closureVoucherNo, setClosureVoucherNo] = useState<string>("");
  const [closureVouchingDate, setClosureVouchingDate] = useState<string>(""); // bound to <input type="datetime-local">

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
      console.warn("[EMD-Closure] Attachments fetch failed", e);
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

  /** ---------- Fetch details ---------- */
  async function fetchEMDDetails(id: number) {
    const spx = await spCrudOps;

    const select =
      "ID,Title,Status,Modified," +
      // Requestor
      "EmployeeName,EmployeeCode,Department,Division,Location,RM,HOD,ContactNo,EmployeeStatus,Email," +
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
      "DateofReceipt,BankAccount,Amount,ARComment," +
      // ⬇️ New Closure fields to prefill bottom inputs
      "ClosureVoucherNo,ClosureVouchingDate,ApprovalMatrix,WFHistory";

    const expand = "VendorName,TenderType,Currency,ModeofPayment,ContractType,ProductType";

    // Try via BAL
    const res = await spx.getData("EMDDetails", select, expand, `ID eq ${id}`, { column: "ID", isAscending: true }, 1, props);
    if (res && res.length) return res[0];

    // Fallback PnP
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
        "ClosureVoucherNo",
        "ClosureVouchingDate",
        "Currency/Currency",
        "Currency/Title",
        "VendorName/Title",
        "VendorName/Name",
        "TenderType/TenderType",
        "TenderType/Title",
        "ModeofPayment/Title",
        "ContractType/Title",
        "ProductType/Title",
        "ApprovalMatrix",
        "WFHistory"
      )
      .expand("VendorName", "TenderType", "Currency", "ModeofPayment", "ContractType", "ProductType")
      .get();
    return pnp;
  }

  /** ---------- Load details ---------- */
  useEffect(() => {
    if (!itemId || Number.isNaN(itemId)) {
      console.warn("[EMD-Closure] Invalid ItemId", itemId);
      return;
    }

    void (async () => {
      const details = await fetchEMDDetails(itemId);
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

      // Vouching details (AP display)
      const vDate = details.VouchingDate ? String(details.VouchingDate).split("T")[0] : "";
      setVouchingDate(vDate);
      setGLCode(details.GLCode || "");
      setVoucherNo(details.VoucherNo || "");
      setAPTeamComment(details.APTeamComment || "");
      const apD = details.VouchingDate || details.Modified;
      setAPActionDate(apD ? String(apD).split("T")[0] : "");

      // Approver/MANAC
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

      // ⬇️ Prefill new closure inputs if already present
      setClosureVoucherNo(details.ClosureVoucherNo || "");
      setClosureVouchingDate(toInputDateTimeLocal(details.ClosureVouchingDate));
      // 🔥 SET APPROVAL MATRIX
      try {
        const matrix =
          typeof details.ApprovalMatrix === "string"
            ? JSON.parse(details.ApprovalMatrix)
            : details.ApprovalMatrix || [];

        // sort by sequence
        matrix.sort((a: any, b: any) => a.Seq - b.Seq);

        setApprovalMatrix(matrix);

      } catch (e) {
        console.error("Error parsing ApprovalMatrix", e);
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

  /** ---------- Validation ---------- */
  const canClose = () => {
    if (!closureVouchingDate) return false; // datetime-local required
    if (!closureVoucherNo.trim()) return false;
    return true;
  };

  /** ---------- Close EMD ---------- */
  // const onCloseEMD = async () => {
  //   if (isSaving) return;
  //   if (!itemId || Number.isNaN(itemId)) {
  //     alert("Invalid item id.");
  //     return;
  //   }
  //   if (!canClose()) {
  //     alert("Please fill Closure Vouching Date & Closure Voucher No.");
  //     return;
  //   }
  //   if (currentStatus === FINAL_STATUS) {
  //     alert("This request is already Closed.");
  //     history.push(ROUTE_AFTER_CLOSE);
  //     return;
  //   }

  //   try {
  //     setIsSaving(true);
  //     const spx = await spCrudOps;

  //     // Convert local datetime to ISO for SharePoint DateTime
  //     const closureVouchingDateISO = toISOStringFromLocal(closureVouchingDate);

  //     // Update ONLY closure fields + final status
  //     await spx.updateData(
  //       "EMDDetails",
  //       itemId,
  //       {
  //         ClosureVouchingDate: closureVouchingDateISO, // DateTime
  //         ClosureVoucherNo: closureVoucherNo,          // Text
  //         Status: FINAL_STATUS,                        // "Closed"
  //       },
  //       props
  //     );

  //     alert("✅ EMD closed successfully.");
  //     history.push(ROUTE_AFTER_CLOSE);
  //   } catch (e) {
  //     console.error("[EMD-Closure] Close EMD error", e);
  //     alert("Something went wrong. Please check console.");
  //   } finally {
  //     setIsSaving(false);
  //   }
  // };

  const onCloseEMD = async () => {

    if (isSaving) return;

    if (!itemId || Number.isNaN(itemId)) {
      alert("Invalid item id.");
      return;
    }

    const validateClosureInputs = (): boolean => {

      //  Date validation
      if (!closureVouchingDate || closureVouchingDate.trim() === "") {
        alert("Closure Vouching Date is required.");
        return false;
      }

      //  Voucher No validation
      if (!closureVoucherNo || closureVoucherNo.trim() === "") {
        alert("Closure Voucher No is required.");
        return false;
      }
      return true;
    };

    if (!validateClosureInputs()) {
      return;
    }

    // if (!canClose()) {
    //   alert("Please fill Closure Vouching Date & Closure Voucher No.");
    //   return;
    // }

    if (currentStatus === FINAL_STATUS) {
      alert("This request is already Closed.");
      history.push(ROUTE_AFTER_CLOSE);
      return;
    }

    try {
      setIsSaving(true);
      const spx = await spCrudOps;

      // 🔥 1. GET CURRENT ITEM WITH MATRIX
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

      // 🔥 3. SORT
      matrix.sort((a: any, b: any) => a.Seq - b.Seq);

      // 🔥 4. FIND CURRENT PENDING
      const currentIndex = matrix.findIndex(x => x.Status === "Pending");

      // 🔥 5. MARK FINAL STEP
      if (currentIndex !== -1) {
        matrix[currentIndex].Status = "Approved";
      }

      // 🔥 OPTIONAL: Ensure all previous are Approved
      matrix = matrix.map((step, index) => {
        if (index <= currentIndex) {
          return { ...step, Status: "Approved" };
        }
        return step;
      });

      // 🔥 6. CONVERT DATE
      const closureVouchingDateISO = toISOStringFromLocal(closureVouchingDate);

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
        props.context?.pageContext?.user?.displayName || "Closure Team";

      const newEntry = {
        CurrentApprover: currentUser,
        ActionTaken: "EMD Closed",
        Comment: closureComments || "Closure completed",
        Date: new Date().toISOString()
      };

      prevHistory.push(newEntry);

      const wfHistoryPayload = JSON.stringify(prevHistory);

      console.log("🔥 FINAL WFHistory:", prevHistory);
      // 🔥 7. UPDATE SHAREPOINT
      await spx.updateData(
        "EMDDetails",
        itemId,
        {
          ClosureVouchingDate: closureVouchingDateISO,
          ClosureVoucherNo: closureVoucherNo,

          // ✅ FINAL STATUS
          Status: FINAL_STATUS, // "Closed"

          // ✅ APPROVAL FLOW COMPLETE
          ApprovalMatrix: JSON.stringify(matrix),
          CurrentApproverId: null,
          PendingAt: "Completed",
          WFHistory: wfHistoryPayload
        },
        props
      );

      console.log("🔥 FINAL MATRIX AFTER CLOSE:", matrix);

      alert("✅ EMD closed successfully.");
      history.push(ROUTE_AFTER_CLOSE);

    } catch (e) {
      console.error("[EMD-Closure] Close EMD error", e);
      alert("Something went wrong. Please check console.");
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="forex-wrapper">
      {/* ================= HEADER ================= */}
      <div className="forex-header">
        <h2>EMD Closure Request</h2>
        {title ? <div style={{ marginTop: 4, color: "#666" }}>Ref: {title}</div> : null}
        {currentStatus ? (
          <div style={{ marginTop: 4, color: "#999", fontSize: 12 }}>
            Current Status: <b>{currentStatus}</b>
          </div>
        ) : null}
      </div>

      <div className="forex-card">

        {/* ================= APPROVAL HIERARCHY ================= */}
        {/* <div className="emd-hierarchy">
          <div className="emd-step active-step">{employee.EmployeeName}</div>

          <div className="emd-step"  style={{marginLeft:"30px"}}>{employee.ReportingManager}</div>

          <div className="emd-step" style={{marginLeft:"30px"}}>{employee.HOD}</div>
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
              const isLast = index === approvalMatrix.length - 1;

              // 🔴 Rejected
              if (step.Status === "Rejected") {
                stepClass = "red";
              }

              // 🟢 Approved
              else if (step.Status === "Approved") {
                stepClass = "green";
              }

              // 🟢 Final Accepted
              else if (isLast && step.Status === "Accepted") {
                stepClass = "green";
              }

              // 🟠 Current Pending
              else if (index === firstPending) {
                stepClass = "green";
              }

              return (
                <div key={index} className={`flowStep ${stepClass}`}>
                  {step.Approver || step.ApproverName || step.Role}
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

        {/* Vouching Details (display) */}
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

        {/* Closure Details (read only) */}
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
            <Field label="Approval By">
              <input value="AR Team" readOnly />
            </Field>
            <Field label="Action Taken">
              <input value="Approved" readOnly />
            </Field>
            <Field label="Action Date">
              <input value={approverActionDate || "-"} readOnly />
            </Field>

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
              <input value="AR Team" readOnly />
            </Field>
            <Field label="Comment">
              <textarea value={approverComment || "Approved"} readOnly rows={3} />
            </Field>
            <Field label="Comment Date">
              <input value={approverActionDate || "-"} readOnly />
            </Field>

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
              <textarea value={"Approving for Vouching"} readOnly rows={3} />
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
                {wfHistory.length > 0 ? (
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
                )}
              </tbody>
            </table>
          </div>
        </Section>
        {/* ===== Bottom Closure Inputs (user-entered) + Close Button ===== */}
        <Section title="Closure Voucher Details">
          <Grid>
            <Field label="Closure Vouching Date*">
              <input
                type="datetime-local"
                value={closureVouchingDate}
                onChange={(e) => setClosureVouchingDate(e.target.value)}
              />
            </Field>
            <Field label="Closure Voucher No*">
              <input
                value={closureVoucherNo}
                onChange={(e) => setClosureVoucherNo(e.target.value)}
                placeholder="Enter Closure Voucher No"
              />
            </Field>
          </Grid>

          <div className="button-row" style={{ marginTop: 12 }}>
            <button
              className="btn-submit"
              // disabled={isSaving || !canClose()}
              onClick={onCloseEMD}
            >
              {isSaving ? "Closing..." : "Close EMD"}
            </button>
            <button className="btn-exit" onClick={() => history.push(ROUTE_AFTER_CLOSE)}>
              Exit
            </button>
          </div>
        </Section>
      </div>
    </div>
  );
};

export default EMDClosureRequestForm;
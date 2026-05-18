import React, { useEffect, useMemo, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory, useLocation } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { sp } from "@pnp/sp";
import logo from "../../assets/sona-comstarlogo.png";

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
  const [productType, setProductType] = useState<string>("");
  const [modeofPayment, setModeofPayment] = useState<string>("");
  // const [modeOfPaymentText, setModeOfPaymentText] = useState<string>("");
  const [contractTypeText, setContractTypeText] = useState<string>("");
  // const [productTypeText, setProductTypeText] = useState<string>("");

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
      "ProductType/ProductType" +
      "ModeofPayment/Mode" +
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
        "ModeofPayment/Id",
        "ModeofPayment/Mode",
        "ProductType/Id",
        "ProductType/ProductType",
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

      setProductType(details.ProductType?.ProductType || "");
      setModeofPayment(details.ModeofPayment?.Mode);

      // Lookups
      setVendorNameTitle(details.VendorName?.Name || "");
      setTenderTypeText(details.TenderType?.Title || details.TenderType?.TenderType || "");
      setCurrencyText(details.Currency?.Title || details.Currency?.Currency || "");
      // setModeOfPaymentText(details.ModeofPayment?.Title || "");
      setContractTypeText(details.ContractType?.Title || "");
      // setProductTypeText(details.ProductType?.Title || "");

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
      const today = new Date();

      const formattedDate =
        String(today.getDate()).padStart(2, '0') + '/' +
        String(today.getMonth() + 1).padStart(2, '0') + '/' +
        today.getFullYear();
      const newEntry = {
        CurrentApprover: currentUser,
        ActionTaken: "EMD Closed",
        Comment: closureComments || "Closure completed",
        Date: formattedDate
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
    <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
      <div className='row'>
        <div className='col-md-12'>
          <div className='Main-Boxpoup'>
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1>EMD Closure Request</h1>
              {title ? <div style={{ marginTop: 4, color: "#666" }}>Ref: {title}</div> : null}
              {currentStatus ? (
                <div style={{ marginTop: 4, color: "#999", fontSize: 12 }}>
                  Current Status: <b>{currentStatus}</b>
                </div>
              ) : null}
            </div>
            <div className="approvalFlow">

              {/* Initiator */}
              <ul >
                <li className="flowStep green">
                  {employee.EmployeeName || "Initiator"}
                </li>
              </ul>

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
                  stepClass = "orange";
                }

                return (
                  <ul key={index}>
                    <li className={`flowStep ${stepClass}`}>
                      {step.Approver || step.ApproverName || step.Role}
                    </li>
                  </ul>
                );
              })}

            </div>
            <div className="borderedbox">
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
                    <label htmlFor="Employee Email" className='font'>Email </label> : &nbsp;&nbsp;
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
                    <input value={vendorCode} className='form-control readonly' readOnly />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>Vendor Name </label>
                    <input value={vendorNameTitle} className='form-control readonly' readOnly />

                  </div>
                  <div className='col-md-4'>
                    <label className='font'>Vendor Site </label>.
                    <input value={vendorSite} readOnly className="form-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  {/* <div className='col-md-4'>
                                    <label className="font">Contract Type </label>
                                    <input value={tenderNo} readOnly className="form-control readonly" />
                                </div> */}
                  <div className='col-md-4'>
                    <label className="font">Tender No </label>
                    <input value={tenderNo} readOnly className="form-control readonly" />

                  </div>
                  <div className='col-md-4'>
                    <label className="font">Tender Date </label>
                    <input type="date" value={tenderDate} readOnly className="form-control readonly" />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Tender Type </label>
                    <input value={tenderTypeText} readOnly className="form-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Tender Amount </label>
                    <input value={formatINR(tenderAmount)} readOnly className="form-control readonly" />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">EMD Amount </label>
                    <input value={formatINR(emdAmount)} readOnly className="form-control readonly" />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Currency </label>
                    <input value={currencyText} readOnly className="form-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Tender Closing Date </label>
                    <input type="date" value={tenderClosingDate} readOnly className="form-control readonly" />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">EMD Percentage </label>
                    <input value={emdPercentage} className="form-control readonly" readOnly />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Mode of Payment </label>
                    <input value={productType} readOnly className="form-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Product Type </label>
                    <input type="text" value={modeofPayment} readOnly className="form-control readonly" />
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
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vouching Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className='font'>Vouching Date </label>
                    <input type="date" className='form-control' value={vouchingDate} readOnly />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>GL </label>
                    <input value={glCode} readOnly className='form-control' />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>vendor Code </label>
                    <input value={vendorCode} readOnly className="form-control" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">voucher No. </label>
                    <input value={voucherNo} readOnly className="form-control" />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>UTR Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className='font'>UTR No </label>
                    <input className='form-control' value={utrNo} readOnly />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>UTR Date </label>
                    <input type="date" value={utrDate} readOnly className='form-control' />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Closure Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className='font'>Date of Receipt <span className="Mantorystar">*</span> </label>
                    <input type="date" className='form-control' value={dateOfReceipt} readOnly />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>Bank Account <span className="Mantorystar">*</span> </label>
                    <input value={bankAccount} readOnly className='form-control' />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>Amount <span className="Mantorystar">*</span> </label>
                    <input value={closureAmount} readOnly className='form-control' />
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className='font'>Comments <span className="Mantorystar">*</span> </label>
                      <textarea rows={3} value={closureComments} readOnly className='form-control' />
                    </div>
                  </div>
                </div>
              </div>
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
                <label>Closure Voucher Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className='font'>Closure Vouching Date <span className="Mantorystar">*</span> </label>
                    <input className="form-control"
                      type="datetime-local"
                      value={closureVouchingDate}
                      onChange={(e) => setClosureVouchingDate(e.target.value)}
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>Closure Voucher No  <span className="Mantorystar">*</span> </label>
                    <input type="text" value={closureVoucherNo} onChange={(e) => setClosureVoucherNo(e.target.value)} className='form-control' />
                  </div>
                </div>
              </div>

              <div className='row my-3'>
                <div className='col-md-12'>
                  <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                    <button
                      className="reject-btn"
                      onClick={onCloseEMD}
                    >
                      {isSaving ? "Closing..." : "Close EMD"}
                    </button>
                    <button className="reset-btn" onClick={() => history.push(ROUTE_AFTER_CLOSE)}>Exit</button>
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

export default EMDClosureRequestForm;
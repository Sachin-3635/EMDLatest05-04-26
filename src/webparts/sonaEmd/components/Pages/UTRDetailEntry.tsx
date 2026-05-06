import React, { useEffect, useMemo, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory, useLocation } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { sp } from "@pnp/sp";
import logo from "../../assets/sona-comstarlogo.png";

// ---------- UI helpers ----------
const Section = ({ title, children }: any) => (
  <div className="form-section">
    <h3>{title}</h3>
    {children}
  </div>
);
const Grid = ({ children }: any) => <div className="form-grid">{children}</div>;
const Field = ({ label, children, full }: any) => (
  <div className={full ? "form-field full" : "form-field"}>
    <label>{label}</label>
    {children}
  </div>
);

const UTRDetailEntry = (props: ISonaEmdProps) => {
  const history = useHistory();
  const location = useLocation();
  const spCrudOps = SPCRUDOPS();

  // ---------- Setup PnP ----------
  useEffect(() => {
    try {
      sp.setup({ spfxContext: (props as any).context });
    } catch {
      // no-op
    }
  }, [props]);

  // ---------- Parse ItemId (search + hash; HashRouter-safe) ----------
  const itemId = useMemo(() => {
    let idStr = new URLSearchParams(location.search || window.location.search).get("ItemId");
    if (!idStr) {
      const hash = window.location.hash || "";
      const qIndex = hash.indexOf("?");
      const search = qIndex >= 0 ? hash.substring(qIndex) : "";
      idStr = new URLSearchParams(search).get("ItemId");
    }
    const parsed = idStr ? parseInt(idStr, 10) : NaN;
    console.log("[UTR] Parsed ItemId:", idStr, "=>", parsed);
    return parsed;
  }, [location.search, location.hash]);

  // ---------- State ----------
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

  // Attachments
  const [attachments, setAttachments] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);

  // Treasury Action (required)
  const [utrNo, setUTRNo] = useState<string>("");
  const [utrDate, setUTRDate] = useState<string>("");
  const [treasuryComment, setTreasuryComment] = useState<string>("");

  // New attachments (Treasury) - currently not used in UI (upload section commented)
  const [files, setFiles] = useState<File[]>([]);

  // Prevent duplicate submit clicks
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [wfHistory, setWfHistory] = useState<any[]>([]);
  // ---------- Helpers (function declarations ARE hoisted) ----------
  async function fetchEMDDetailsRobust(id: number) {
    const spx = await spCrudOps;

    // No ApproverActionDate to avoid 400s if not present
    const select =
      "ID,Title,Status,Modified," +
      // Requestor
      "EmployeeName,EmployeeCode,Department,Division,Location,RM,HOD,ContactNo,EmployeeStatus,Email," +
      // EMD details
      "VendorCode,VendorSite," +
      "TenderNo,TenderDate,TenderAmount,EMDAmount,EMDPercentage,TenderClosingDate," +
      "TenderTypeId,TenderType/TenderType," +
      "VendorNameId,VendorName/Title," +
      "CurrencyId,Currency/Currency,ApprovalMatrix,WfHistory" +
      // Vouching (AP)
      "VouchingDate,GLCode,VoucherNo,APTeamComment," +
      // Approver
      "ApproverComment,WFHistory" +
      // Treasury
      "UTRNo,UTRDate,TreasuryComment";

    const expand = "VendorName,TenderType,Currency";

    const tries = [
      { filter: `ID eq ${id}`, order: { column: "ID", isAscending: true } },
      { filter: `Id eq ${id}`, order: { column: "Id", isAscending: true } },
    ];

    for (const t of tries) {
      try {
        console.log("[UTR] getData with", t.filter);
        const res = await spx.getData("EMDDetails", select, expand, t.filter, t.order, 1, props);
        if (res && res.length) return res[0];
      } catch (e) {
        console.warn("[UTR] getData error with", t.filter, e);
      }
    }

    // Fallback PnP
    try {
      console.log("[UTR] Fallback PnP getById");
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
          "Currency/Currency",
          "VendorName/Name",
          "VendorName/Title",
          "TenderType/TenderType",
          "ApprovalMatrix",
          "WFHistory"
        )
        .expand("VendorName", "TenderType", "Currency")
        .get();
      return pnp;
    } catch (e) {
      console.error("[UTR] PnP fallback error", e);
      return null as any;
    }
  }

  async function loadAttachments(id: number) {
    try {
      const files = await sp.web.lists.getByTitle("EMDDetails").items.getById(id).attachmentFiles();
      console.log("[UTR] Attachments:", files);
      setAttachments(files || []);
    } catch (e) {
      console.warn("[UTR] Attachments fetch failed", e);
      setAttachments([]);
    }
  }

  // ---------- Load details ----------
  useEffect(() => {
    if (!itemId || Number.isNaN(itemId)) {
      console.warn("[UTR] Invalid ItemId", itemId);
      return;
    }

    // mark the promise as intentionally not awaited to satisfy no-floating-promises
    void (async () => {
      const details = await fetchEMDDetailsRobust(itemId);
      if (!details) {
        alert("No data found for this ItemId.");
        return;
      }

      console.log("[UTR] Details:", details);

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
      // Derived %
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
      const appD = details.Modified; // If ApproverActionDate added later, prefer that.
      setApproverActionDate(appD ? String(appD).split("T")[0] : "");

      // Treasury prefill (if already set)
      setUTRNo(details.UTRNo || "");
      setUTRDate(details.UTRDate ? String(details.UTRDate).split("T")[0] : "");
      setTreasuryComment(details.TreasuryComment || "");
      // 🔥 Approval Matrix Load
      try {
        let matrix: any[] = [];

        if (details?.ApprovalMatrix) {
          matrix =
            typeof details.ApprovalMatrix === "string"
              ? JSON.parse(details.ApprovalMatrix)
              : details.ApprovalMatrix;
        }

        if (Array.isArray(matrix)) {
          matrix.sort((a: any, b: any) => a.Seq - b.Seq);
          setApprovalMatrix(matrix);
        }

        console.log("ApprovalMatrix 👉", matrix);
        // 🔥 WF HISTORY LOAD
        try {
          let history: any[] = [];

          if (details?.WFHistory) {
            history =
              typeof details.WFHistory === "string"
                ? JSON.parse(details.WFHistory)
                : details.WFHistory;
          }

          setWfHistory(history || []);

          console.log("WFHistory 👉", history);

        } catch (e) {
          console.error("WFHistory parse error", e);
        }

      } catch (e) {
        console.error("ApprovalMatrix parse error", e);
      }
      // Attachments
      await loadAttachments(itemId);
    })();
  }, [itemId]);

  // ---------- Submit ----------
  // const onsubmit = async () => {
  //   if (isSaving) return;

  //   try {
  //     setIsSaving(true);

  //     if (!itemId || Number.isNaN(itemId)) return alert("Invalid item id.");
  //     if (!utrNo.trim()) return alert("UTR No. is required.");
  //     if (!utrDate) return alert("UTR Date is required.");
  //     if (!treasuryComment.trim()) return alert("Comments are required.");

  //     const spx = await spCrudOps;

  //     // 👇 Your required next status
  //     const NEXT_STATUS = "EMD Paid";

  //     // Optional: if already paid, just route ahead
  //     if (currentStatus === NEXT_STATUS) {
  //       alert("Already marked as EMD Paid.");
  //       history.push("/EMDClousureDashboard");
  //       return;
  //     }

  //     await spx.updateData(
  //       "EMDDetails",
  //       itemId,
  //       {
  //         UTRNo: utrNo.trim(),
  //         UTRDate: utrDate,
  //         TreasuryComment: treasuryComment.trim(),
  //         Status: NEXT_STATUS, // 👈 set status here
  //       },
  //       props
  //     );

  //     if (files.length > 0) {
  //       await Promise.all(
  //         files.map((file) => spx.addAttchmentInList(file, "EMDDetails", itemId, file.name, props))
  //       );
  //     }

  //     alert("✅ EMD marked as Paid and UTR details saved.");
  //     history.push("/EMDClousureDashboard"); // 👈 navigate to EMD Closure Dashboard
  //   } catch (e) {
  //     console.error("[UTR] Submit error", e);
  //     alert("Something went wrong. Please check console.");
  //   } finally {
  //     setIsSaving(false);
  //   }
  // };

  const getUpdatedApprovalMatrix = async (oldMatrix: any[]) => {

    const spx = await spCrudOps;

    // 🔥 1. GET MASTER APPROVAL MATRIX
    const master = await spx.getData(
      "EMDApprovalMatrix",
      "ID,Role/RoleName,Approver/ID,Approver/Title",
      "Role,Approver",
      "RequestType eq 'EMD Approval'",
      { column: "ID", isAscending: true },
      1000,
      props
    );

    console.log("📊 MASTER MATRIX 👉", master);

    // 🔥 2. MAP MASTER
    const masterMatrix = master.map((x: any, index: number) => ({
      Seq: index + 1,
      Role: x.Role?.RoleName,
      ApproverID: x.Approver?.ID,
      Approver: x.Approver?.Title,
    }));

    // 🔥 3. MERGE OLD + MASTER
    const updatedMatrix = masterMatrix.map((newItem: { Role: any; ApproverID: any; Approver: any; }) => {

      const oldItem = oldMatrix.find(
        (x) => x.Role === newItem.Role   // 🔥 match by Role
      );

      if (!oldItem) {
        // new role added
        return {
          ...newItem,
          Status: "Pending"
        };
      }

      // 🔥 CHECK IF APPROVER CHANGED
      const isChanged =
        String(oldItem.ApproverID) !== String(newItem.ApproverID);

      if (isChanged) {
        console.log(`⚠️ Approver changed for ${newItem.Role}`);
        console.log("OLD 👉", oldItem.Approver, oldItem.ApproverID);
        console.log("NEW 👉", newItem.Approver, newItem.ApproverID);
      }

      return {
        ...newItem,
        Status: oldItem.Status   // ✅ preserve status
      };
    });

    console.log("✅ FINAL MERGED MATRIX 👉", updatedMatrix);

    return updatedMatrix;
  };


  //old where approveer matrix issue come
  const onsubmit = async () => {

    if (isSaving) return;

    try {
      setIsSaving(true);

      if (!itemId || Number.isNaN(itemId)) return alert("Invalid item id.");
      if (!utrNo.trim()) return alert("UTR No. is required.");
      if (!utrDate) return alert("UTR Date is required.");
      if (!treasuryComment.trim()) return alert("Comments are required.");

      const spx = await spCrudOps;

      // 🔥 1. GET CURRENT ITEM
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

      // 🔥 2. PARSE OLD MATRIX
      let oldMatrix: any[] = [];

      try {
        oldMatrix =
          typeof item.ApprovalMatrix === "string"
            ? JSON.parse(item.ApprovalMatrix)
            : item.ApprovalMatrix || [];
      } catch {
        oldMatrix = [];
      }

      if (!Array.isArray(oldMatrix) || oldMatrix.length === 0) {
        alert("Approval matrix empty");
        return;
      }

      // 🔥 3. GET MASTER MATRIX
      const master = await spx.getData(
        "EMDApprovalMatrix",
        "ID,Role/RoleName,Approver/ID,Approver/Title",
        "Role,Approver",
        "RequestType eq 'EMD Approval'",
        { column: "ID", isAscending: true },
        1000,
        props
      );

      const masterMatrix = master.map((x: any, index: number) => ({
        Seq: index + 1,
        Role: x.Role?.RoleName,
        ApproverID: x.Approver?.ID,
        Approver: x.Approver?.Title,
      }));

      // 🔥 4. MERGE OLD + MASTER (VERY IMPORTANT)
      let matrix = masterMatrix.map((newItem: { Role: any; ApproverID: any; }) => {

        const oldItem = oldMatrix.find(
          (x) => x.Role === newItem.Role
        );

        if (!oldItem) {
          return {
            ...newItem,
            Status: "Not Started"
          };
        }

        const isChanged =
          String(oldItem.ApproverID) !== String(newItem.ApproverID);

        // 🔥 IF APPROVER CHANGED → RESET
        if (isChanged) {
          console.log(`⚠️ Approver changed for ${newItem.Role}`);
          return {
            ...newItem,
            Status: "Not Started"
          };
        }

        return {
          ...newItem,
          Status: oldItem.Status
        };
      });

      // 🔥 5. ADD ROLES NOT IN MASTER (RM/HOD SAFE)
      oldMatrix.forEach(oldItem => {
        const exists = matrix.find((x: { Role: any; }) => x.Role === oldItem.Role);
        if (!exists) {
          matrix.push(oldItem);
        }
      });

      // 🔥 6. REASSIGN SEQ IN PROPER ORDER (FIX FOR DUPLICATE SEQ)
      let finalMatrix: any[] = [];
      let seq = 1;

      // First, add RM if exists
      const rm = matrix.find((x: any) => x.Role === 'RM' || x.Role === 'Reporting Manager');
      if (rm) {
        finalMatrix.push({ ...rm, Seq: seq++ });
      }

      // Then HOD
      const hod = matrix.find((x: any) => x.Role === 'HOD' || x.Role === 'Head of Department');
      if (hod) {
        finalMatrix.push({ ...hod, Seq: seq++ });
      }

      // Then the rest in their current order
      matrix.forEach((item: any) => {
        if (item !== rm && item !== hod) {
          finalMatrix.push({ ...item, Seq: seq++ });
        }
      });

      matrix = finalMatrix;

      // 🔥 7. ENSURE SINGLE PENDING
      let foundPending = false;

      matrix = matrix.map((item: { Status: string; }) => {

        if (item.Status === "Pending" && !foundPending) {
          foundPending = true;
          return item;
        }

        if (item.Status === "Pending" && foundPending) {
          return {
            ...item,
            Status: "Not Started"
          };
        }

        return item;
      });

      // 🔥 IF NO PENDING → SET FIRST
      if (!matrix.some((x: { Status: string; }) => x.Status === "Pending")) {
        const firstIndex = matrix.findIndex((x: { Status: string; }) => x.Status !== "Approved");
        if (firstIndex !== -1) {
          matrix[firstIndex].Status = "Pending";
        }
      }

      // 🔥 8. FIND CURRENT PENDING
      const currentIndex = matrix.findIndex((x: { Status: string; }) => x.Status === "Pending");

      if (currentIndex === -1) {
        alert("No pending approver");
        return;
      }

      console.log("👉 CURRENT:", matrix[currentIndex]);

      // 🔥 9. APPROVE CURRENT
      matrix[currentIndex] = {
        ...matrix[currentIndex],
        Status: "Approved"
      };

      const nextIndex = currentIndex + 1;

      let payload: any = {};

      // 🔥 10. NEXT APPROVER
      if (nextIndex < matrix.length) {

        const next = matrix[nextIndex];

        // ⭐ VERY IMPORTANT
        matrix[nextIndex] = {
          ...next,
          Status: "Pending"
        };

        payload.CurrentApproverId = Number(next.ApproverID);
        payload.Status = "EMD Paid";
        payload.PendingAt = `Pending at ${next.Approver}`;

        console.log("➡ NEXT:", next);

      } else {

        payload.CurrentApproverId = null;
        payload.Status = "EMD Paid";
        payload.PendingAt = "EMD Paid";

        console.log("✅ FINAL APPROVAL");
      }

      // 🔥 11. UPDATE DATA
      payload.ApprovalMatrix = JSON.stringify(matrix);
      payload.UTRNo = utrNo.trim();
      payload.UTRDate = utrDate;
      payload.TreasuryComment = treasuryComment.trim();

      console.table(matrix);
      let prevHistory: any[] = [];

      try {
        prevHistory =
          typeof item.WFHistory === "string"
            ? JSON.parse(item.WFHistory)
            : item.WFHistory || [];
      } catch {
        prevHistory = [];
      }

      const today = new Date();

      const formattedDate =
        String(today.getDate()).padStart(2, '0') + '/' +
        String(today.getMonth() + 1).padStart(2, '0') + '/' +
        today.getFullYear();

      const newEntry = {
        CurrentApprover: "Treasury",
        ActionTaken: "EMD Paid",
        Comment: treasuryComment,
        Date: formattedDate,
      };

      prevHistory.push(newEntry);

      payload.WFHistory = JSON.stringify(prevHistory);

      // 🔥 12. SAVE
      await spx.updateData("EMDDetails", itemId, payload, props);

      alert("✅ Approved & Updated Successfully");

      // history.push("/EMDClousureDashboard");
      history.goBack();

    } catch (e) {
      console.error("[UTR] Submit error", e);
      alert("Something went wrong.");
    } finally {
      setIsSaving(false);
    }
  };









  // ---------- Formatter ----------

  const formatINR = (val?: string) => {
    if (!val) return "";
    const n = Number(val);
    if (isNaN(n)) return val;
    return n.toLocaleString("en-IN");
  };

  return (
    <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
      <div className='row'>
        <div className='col-md-12'>
          <div className='Main-Boxpoup'>
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1>UTR Detail Entry </h1>
            </div>
            <div className="emd-hierarchy">

              {/* Initiator */}
              <div className="emd-step green">{employee.EmployeeName}</div>

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
                  stepClass = "orange";
                }

                return (
                  <div key={index} className={`emd-step ${stepClass}`}>
                    {step.Approver || step.ApproverName}
                  </div>
                );
              })}

            </div>
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
            <div className="heading1" style={{ marginTop: "10px" }}>
              <label>EMD Request Details</label>
            </div>
            <div className='main-formcontainer'>
              <div className='row mb-20'>
                <div className='col-md-4'>
                  <label className='font'>Vendor Code </label>
                  <input value={vendorCode} className='form-control' readOnly />
                </div>
                <div className='col-md-4'>
                  <label className='font'>Vendor Name </label>
                  <input value={vendorNameTitle} className='form-control' readOnly />

                </div>
                <div className='col-md-4'>
                  <label className='font'>Vendor Site </label>.
                  <input value={vendorSite} readOnly className="form-control" />
                </div>
              </div>
              <div className='row mb-20'>
                <div className='col-md-4'>
                  <label className="font">Contract Type </label>
                  <input value={tenderNo} readOnly className="form-control" />
                </div>
                <div className='col-md-4'>
                  <label className="font">Tender No </label>
                  <input value={tenderNo} readOnly className="form-control" />

                </div>
                <div className='col-md-4'>
                  <label className="font">Tender Date </label>
                  <input type="date" value={tenderDate} readOnly className="form-control" />
                </div>
              </div>
              <div className='row mb-20'>
                <div className='col-md-4'>
                  <label className="font">Tender Type </label>
                  <input value={tenderTypeText} readOnly className="form-control" />
                </div>
                <div className='col-md-4'>
                  <label className="font">Tender Amount </label>
                  <input value={formatINR(tenderAmount)} readOnly className="form-control" />
                </div>
                <div className='col-md-4'>
                  <label className="font">EMD Amount </label>
                  <input value={formatINR(emdAmount)} readOnly className="form-control" />
                </div>

              </div>

              <div className='row mb-20'>
                <div className='col-md-4'>
                  <label className="font">Currency </label>
                  <input value={currencyText} readOnly className="form-control" />
                </div>
                <div className='col-md-4'>
                  <label className="font">Tender Closing Date </label>
                  <input type="date" value={tenderClosingDate} readOnly className="form-control" />
                </div>
                <div className='col-md-4'>
                  <label className="font">EMD Percentage </label>
                  <input value={emdPercentage} className="form-control" readOnly />
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
              <label>Action</label>
            </div>
            <div className='main-formcontainer'>
              <div className='row mb-20'>
                <div className='col-md-4'>
                  <label className='font'> UTR No. <span className='Mantorystar'>*</span></label>
                  <input className="form-control" value={utrNo} onChange={(e) => setUTRNo(e.target.value)} placeholder="Enter UTR No." />
                </div>
                <div className='col-md-4'>
                  <label className='font'> UTR Date. <span className='Mantorystar'>*</span></label>
                  <input type="date" className="form-control" value={utrDate} onChange={(e) => setUTRDate(e.target.value)} />
                </div>
                <div className='col-md-4'>
                  <label className='font'>Comments <span className='Mantorystar'>*</span></label>
                  <textarea
                    rows={3}
                    value={treasuryComment}
                    onChange={(e) => setTreasuryComment(e.target.value)}
                    placeholder="Enter comments" className="form-control"
                  />
                </div>
              </div>

            </div>
            <div className="heading1" style={{ marginTop: "10px" }}>
              <label>Upload Documents</label>
            </div>
            <div className='main-formcontainer'>
              <div className='row mb-20'>
                <div className='col-md-4'>
                  <label className='font'> Attach </label>
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
            <div className='row my-3'>
              <div className='col-md-12'>
                <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                  <button className="submit-btn" disabled={isSaving} onClick={onsubmit}>
                    {isSaving ? "Saving..." : "Paid"}
                  </button>
                  <button className="reset-btn" onClick={() => history.push("/TreasuryLandingPage")}>Exit</button>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default UTRDetailEntry;
import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory } from 'react-router-dom';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import SPCRUDOPS from "../../service/BAL/spcrud";
import { sp } from "@pnp/sp";
import logo from "../../assets/sona-comstarlogo.png";

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

const MANACApprovalForm = (props: ISonaEmdProps) => {
  // -------- Parse ItemId from hash (e.g., #/MANACApprovalForm?ItemId=123) --------
  const hash = window.location.hash || "";
  const qs = new URLSearchParams(hash.split("?")[1] || "");
  const Id = qs.get("ItemId");                   // string | null
  const itemId = Id ? parseInt(Id, 10) : NaN;    // number

  const history = useHistory();
  const spCrudOps = SPCRUDOPS();

  // -------- Dropdown options --------
  const [contractTypeOptions, setContractTypeOptions] = useState<IDropdownOption[]>([]);
  const [modeOfPaymentOptions, setModeOfPaymentOptions] = useState<IDropdownOption[]>([]);
  const [vendorNameOptions, setVendorNameOptions] = useState<IDropdownOption[]>([]);
  const [vendorCodeOptions, setVendorCodeOptions] = useState<IDropdownOption[]>([]);
  const [vendorSiteOptions, setVendorSiteOptions] = useState<IDropdownOption[]>([]);
  const [productTypeOptions, setProductTypeOptions] = useState<IDropdownOption[]>([]);
  const [currencyOptions, setCurrencyOptions] = useState<IDropdownOption[]>([]);

  // -------- Selected keys --------
  const [contractType, setContractType] = useState<number | undefined>();
  const [modeOfPayment, setModeOfPayment] = useState<number | undefined>();
  const [vendorName, setVendorName] = useState<number | undefined>();
  const [vendorCodeKey, setVendorCodeKey] = useState<number | undefined>();
  const [vendorSiteKey, setVendorSiteKey] = useState<number | undefined>();
  const [productType, setProductType] = useState<number | undefined>();
  const [currency, setCurrency] = useState<number | undefined>();

  // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&
  const [isTenderDuplicate, setIsTenderDuplicate] = useState<boolean>(false);
  // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&

  // -------- Texts --------
  const [vendorCode, setVendorCode] = useState("");
  const [vendorSite, setVendorSite] = useState("");
  const [approverMatrix, setApproverMatrix] = useState<any[]>([]);
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
    HODId: 0
  });

  const [vendor, setVendor] = useState({
    TenderNo: "",
    TenderDate: "",
    TenderAmount: "",
    EMDAmount: "",
    TenderClosingDate: "",
    EMDPercentage: "",
    Comments: "",                 // Approver comments (bind to ApproverComment)
  });

  const [status, setStatus] = useState<string>(""); // current status (optional display)
  const [title, setTitle] = useState<string>("");   // EMD Request No. (Title)
  const [actionLoading, setActionLoading] = useState<boolean>(false);
  const [wfHistory, setWfHistory] = useState<any[]>([]);
  const [commentHistory, setCommentHistory] = useState<any[]>([]);

  const [attachments, setAttachments] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);

  // Stage-wise grouped attachments for Supporting Docs
  const [apDocs, setApDocs] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);
  const [treasuryDocs, setTreasuryDocs] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);
  const [manacDocs, setManacDocs] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);

  // -------- Constants for status updates --------
  const APPROVE_STATUS = "Pending for Vouching";
  // If you have a "Rejected" choice in list, use "Rejected" here:
  const REJECT_STATUS = "Pending for Approval"; // or "Rejected"

  // -------- Load masters on mount --------
  useEffect(() => {
    getCurrencyData();
    getVendorDataMaster();
    getContractTypeData();
    getModeOfPaymentData();
    getProductTypeData();
  }, []);

  // -------- After options ready, fetch the item --------
  useEffect(() => {
    if (!itemId || Number.isNaN(itemId)) return;
    if (!vendorCodeOptions.length || !vendorSiteOptions.length) return;

    fetchEMDDetails(String(itemId)).then(details => {
      if (!details) return;

      setTitle(details.Title || "");
      setStatus(details.Status || "Pending for Approval");

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
        HODId: details.HODId || 0
      });

      setVendor({
        TenderNo: details.TenderNo || "",
        TenderDate: details.TenderDate ? details.TenderDate.split("T")[0] : "",
        TenderAmount: details.TenderAmount || "",
        EMDAmount: details.EMDAmount || "",
        TenderClosingDate: details.TenderClosingDate ? details.TenderClosingDate.split("T")[0] : "",
        EMDPercentage: details.EMDPercentage || "",
        Comments: details.ApproverComment || "", // prefill existing comment if any
      });

      // Lookup dropdowns
      setVendorName(details.VendorNameId || undefined);
      setContractType(details.ContractTypeId || undefined);
      setModeOfPayment(details.ModeofPaymentId || undefined);
      setProductType(details.ProductTypeId || undefined);
      setCurrency(details.CurrencyId || undefined);

      // Vendor Code
      const codeOption = vendorCodeOptions.find(opt => opt.text === details.VendorCode);
      setVendorCode(details.VendorCode || "");
      setVendorCodeKey(codeOption?.key as number | undefined);

      // Vendor Site
      const siteOption = vendorSiteOptions.find(opt => opt.text === details.VendorSite);
      setVendorSite(details.VendorSite || "");
      setVendorSiteKey(siteOption?.key as number | undefined);
      // 🔥 Parse ApprovalMatrix from list
      let matrix: any[] = [];

      try {
        matrix =
          typeof details.ApprovalMatrix === "string"
            ? JSON.parse(details.ApprovalMatrix)
            : details.ApprovalMatrix || [];
      } catch {
        matrix = [];
      }

      // sort by sequence
      matrix.sort((a: any, b: any) => a.Seq - b.Seq);
      setApproverMatrix(matrix);

      // 🔥 LOAD WF + COMMENT HISTORY
      try {
        const wf = details.WFHistory
          ? typeof details.WFHistory === "string"
            ? JSON.parse(details.WFHistory)
            : details.WFHistory
          : [];

        const ch = details.CommentHistory
          ? typeof details.CommentHistory === "string"
            ? JSON.parse(details.CommentHistory)
            : details.CommentHistory
          : [];

        setWfHistory(Array.isArray(wf) ? wf : []);
        setCommentHistory(Array.isArray(ch) ? ch : []);

      } catch (e) {
        console.error("History parse error", e);
      }
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [itemId, vendorCodeOptions.length, vendorSiteOptions.length]);

  // Auto calc %
  useEffect(() => {
    const tender = parseFloat(vendor.TenderAmount || "0");
    const emd = parseFloat(vendor.EMDAmount || "0");
    if (tender > 0) {
      const percent = ((emd / tender) * 100).toFixed(2);
      setVendor(prev => ({ ...prev, EMDPercentage: percent }));
    }
  }, [vendor.TenderAmount, vendor.EMDAmount]);

  // -------- Masters --------
  const getCurrencyData = async () => {
    try {
      const sp = await spCrudOps;
      const data = await sp.getData(
        "CurrencyMaster",
        "Title,Id,Currency,Status",
        "",
        "Status eq 'Active'",
        { column: "Id", isAscending: true },
        5000,
        props
      );
      setCurrencyOptions(data.map((m: any) => ({ key: m.Id, text: m.Currency })));
    } catch (error) {
      console.error("Error fetching currency data:", error);
    }
  };

  const getVendorDataMaster = async () => {
    const sp = await spCrudOps;
    const res = await sp.getData(
      "VendorMaster",
      "ID,VendorCode,Name,VendorSite",
      "",
      "",
      { column: "ID", isAscending: true },
      5000,
      props
    );
    setVendorNameOptions(res.map((v: any) => ({ key: v.ID, text: v.Name })));
    setVendorCodeOptions(res.map((v: any) => ({ key: v.ID, text: v.VendorCode })));
    setVendorSiteOptions(res.map((v: any) => ({ key: v.ID, text: v.VendorSite })));
  };

  const getModeOfPaymentData = async () => {
    const sp = await spCrudOps;
    const res = await sp.getData(
      "ModeofPaymentMaster",
      "Title,Id,Mode,Status",
      "",
      "Status eq 'Active'",
      { column: "Title", isAscending: true },
      5000,
      props
    );
    setModeOfPaymentOptions(res.map((m: any) => ({ key: m.Id, text: m.Mode })));
  };

  const getProductTypeData = async () => {
    const sp = await spCrudOps;
    const res = await sp.getData(
      "ProductTypeMaster",
      "Title,Id,ProductType,Status",
      "",
      "Status eq 'Active'",
      { column: "Id", isAscending: true },
      5000,
      props
    );
    setProductTypeOptions(res.map((m: any) => ({ key: m.Id, text: m.ProductType })));
  };

  const getContractTypeData = async () => {
    const sp = await spCrudOps;
    const res = await sp.getData(
      "ContractTypeMaster",
      "Title,Id,ContractType,Status",
      "",
      "Status eq 'Active'",
      { column: "Id", isAscending: true },
      5000,
      props
    );
    setContractTypeOptions(res.map((c: any) => ({ key: c.Id, text: c.ContractType })));
  };


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

  useEffect(() => {
    if (itemId && !Number.isNaN(itemId)) {
      loadAttachments(itemId);
    }
  }, [itemId]);



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
      if (vendor.TenderNo) {
        const isDuplicate = await checkDuplicateTendorNumber(vendor.TenderNo);
        setIsTenderDuplicate(isDuplicate);
      }
    };

    checkDuplicate();
  }, [vendor.TenderNo]);

  // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&


  // -------- Fetch EMD Details --------
  const fetchEMDDetails = async (id: string) => {
    const sp = await spCrudOps;
    const res = await sp.getData(
      "EMDDetails",
      [
        "*",
        "Id",
        "Title",
        "Status",
        "ApproverComment",
        "EmployeeName,EmployeeCode,Division,Department,Location,RM,HOD,ContactNo,EmployeeStatus,Email",
        "VendorCode,VendorSite",
        "TenderTypeId,TenderType/TenderType",
        "VendorNameId,VendorName/Title",
        "ContractTypeId,ContractType/Title",
        "ModeofPaymentId,ModeofPayment/Title",
        "ProductTypeId,ProductType/Title",
        "TenderNo,TenderDate,TenderAmount,EMDAmount,CurrencyId,Currency/Currency,TenderClosingDate,EMDPercentage,WFHistory,CommentHistory"
      ].join(","),
      "VendorName,ContractType,ProductType,Currency,ModeofPayment,TenderType",
      `Id eq ${id}`,
      { column: "Id", isAscending: true },
      1,
      props
    );

    return res && res.length ? res[0] : null;

  };

  // -------- Approve / Reject handlers --------
  const updateStatusWithComment = async (newStatus: string) => {
    if (!itemId || Number.isNaN(itemId)) {
      alert("Invalid item id.");
      return;
    }
    const sp = await spCrudOps;

    setActionLoading(true);
    try {
      await sp.updateData(
        "EMDDetails",
        itemId,
        {
          Status: newStatus,                          // Choice column
          ApproverComment: vendor.Comments || "",     // Text column internal name
          // (Optional) add audit fields if you have them:
          // ApprovedBy: props.userDisplayName,
          // ApprovedOn: new Date().toISOString(),
        },
        props
      );

      alert(`✅ Request ${newStatus === APPROVE_STATUS ? "approved" : "updated"} successfully!`);
      history.push("/ApprovalDashboard");
    } catch (err) {
      console.error("Status update error:", err);
      alert("❌ Failed to update status. Please check console.");
    } finally {
      setActionLoading(false);
    }
  };

  const getUpdatedApprovalMatrix = async (existingMatrix: any[]) => {

    const sp = await spCrudOps;

    // 🔥 GET LATEST APPROVERS FROM MASTER
    const master = await sp.getData(
      "EMDApprovalMatrix",
      "ID,Role/RoleName,Approver/ID,Approver/Title",
      "Role,Approver",
      "RequestType eq 'EMD Approval'",
      { column: "ID", isAscending: true },
      5000,
      props
    );

    // 🔥 CREATE MASTER MATRIX
    const masterMatrix = master.map((m: any, index: number) => ({
      Seq: index + 1,
      Role: m.Role?.RoleName,
      ApproverID: m.Approver?.ID,
      Approver: m.Approver?.Title
    }));

    // 🔥 MERGE OLD + NEW (IMPORTANT)
    const updatedMatrix = existingMatrix.map((oldItem) => {

      // find matching role in master
      const newItem = masterMatrix.find((m: { Role: any; }) => m.Role === oldItem.Role);

      // ❗ IF NOT FOUND (RM/HOD etc) → KEEP OLD
      if (!newItem) {
        return oldItem;
      }

      // ❗ CHECK IF APPROVER CHANGED
      const isChanged = oldItem.ApproverID != newItem.ApproverID;

      return {
        ...oldItem,
        ApproverID: isChanged ? newItem.ApproverID : oldItem.ApproverID,
        Approver: isChanged ? newItem.Approver : oldItem.Approver,
        Status: isChanged ? "Pending" : oldItem.Status // reset if changed
      };
    });

    return updatedMatrix;
  };

  const handleApprove = async () => {

    if (!itemId || Number.isNaN(itemId)) {
      alert("Invalid item id.");
      return;
    }

    const sp = await spCrudOps;
    setActionLoading(true);

    try {

      // 🔥 1. GET ITEM
      const items = await sp.getData(
        "EMDDetails",
        "Id,ApprovalMatrix",
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

      // 🔥 3. CURRENT INDEX
      const currentIndex = oldMatrix.findIndex(x => x.Status === "Pending");

      if (currentIndex === -1) {
        alert("No pending approver");
        return;
      }

      // 🔥 4. MERGE MATRIX (DO NOT CHANGE STATUS HERE)
      let newMatrix = await getUpdatedApprovalMatrix(oldMatrix);

      // 🔥 5. APPLY CORRECT STATUS FLOW
      for (let i = 0; i < newMatrix.length; i++) {

        if (i < currentIndex) {
          newMatrix[i].Status = "Approved";
        }

        else if (i === currentIndex) {
          newMatrix[i].Status = "Approved";
        }

        else if (i === currentIndex + 1) {
          newMatrix[i].Status = "Pending";
        }

        else {
          newMatrix[i].Status = "Not Started";
        }
      }

      const nextIndex = currentIndex + 1;

      let payload: any = {};
      payload.ApprovalMatrix = JSON.stringify(newMatrix);

      const isLastApproval = currentIndex === newMatrix.length - 1;

      // 🔥 6. STATUS LOGIC WITH CFO CONDITION
      if (isLastApproval) {

        // ✅ Last approver (CFO already approved)
        payload.CurrentApproverId = null;
        payload.Status = "Completed";
        payload.PendingAt = "Completed";

      } else {

        const next = newMatrix[nextIndex];

        payload.CurrentApproverId = Number(next.ApproverID);

        // 🔥 KEY CHANGE: CFO CHECK
        if (next.Role === "AP Performer") {

          payload.Status = "Pending for Vouching";   // ✅ YOUR REQUIREMENT
          payload.PendingAt = "AP Performer";

        } else {

          payload.Status = "Pending for Approval";
          payload.PendingAt = `Pending at ${next.Role}`;
        }
      }

      payload.ApproverComment = vendor.Comments || "";

      // 🧪 DEBUG (REMOVE AFTER TEST)
      console.log("UPDATED MATRIX:", newMatrix);
      // 🔥 HISTORY UPDATE
      const today = new Date();

      const formattedDate =
        String(today.getDate()).padStart(2, '0') + '/' +
        String(today.getMonth() + 1).padStart(2, '0') + '/' +
        today.getFullYear();

      const currentUser =
        props.context.pageContext.user.displayName ||
        props.context.pageContext.user.email;

      let prevHistory: any[] = [];
      let prevComments: any[] = [];

      try {
        prevHistory = wfHistory || [];
        prevComments = commentHistory || [];
      } catch {
        prevHistory = [];
        prevComments = [];
      }

      const actionTaken = isLastApproval ? "Completed" : "Approved";

      const historyEntry = {
        CurrentApprover: currentUser,
        ActionTaken: actionTaken,
        Comment: vendor.Comments || "",
        Date: formattedDate,
        CurrentStatus: payload.PendingAt || ""
      };

      // 🔥 PUSH
      prevHistory.push(historyEntry);
      prevComments.push(historyEntry);

      // 🔥 SAVE IN PAYLOAD
      payload.WFHistory = JSON.stringify(prevHistory);
      payload.CommentHistory = JSON.stringify(prevComments);

      // 🔥 7. UPDATE
      await sp.updateData("EMDDetails", itemId, payload, props);
      setVendor(prev => ({ ...prev, Comments: "" }));
      alert("✅ Approved Successfully");
      history.push("/ApprovalDashboard");

    } catch (err) {
      console.error("Approval Error:", err);
      alert("Error during approval");
    } finally {
      setActionLoading(false);
    }
  };
  // const handleApprove = async () => {
  //   await updateStatusWithComment(APPROVE_STATUS);
  // };

  // const handleReject = async () => {
  //   await updateStatusWithComment(REJECT_STATUS);
  // };
  // const handleReject = async () => {

  //   // ✅ Comments validation
  //   if (!vendor.Comments || vendor.Comments.trim() === "") {
  //     alert("Comments are required for rejection.");
  //     return;
  //   }

  //   const sp = await spCrudOps;

  //   try {

  //     const items = await sp.getData(
  //       "EMDDetails",
  //       "Id,ApprovalMatrix",
  //       "",
  //       `Id eq ${itemId}`,
  //       { column: "Id", isAscending: false },
  //       1,
  //       props
  //     );

  //     let matrix =
  //       typeof items[0].ApprovalMatrix === "string"
  //         ? JSON.parse(items[0].ApprovalMatrix)
  //         : items[0].ApprovalMatrix;

  //     const currentIndex = matrix.findIndex((x: { Status: string; }) => x.Status === "Pending");

  //     if (currentIndex !== -1) {
  //       matrix[currentIndex].Status = "Rejected";
  //     }

  //     await sp.updateData("EMDDetails", itemId, {
  //       ApprovalMatrix: JSON.stringify(matrix),
  //       Status: "Rejected",
  //       PendingAt: "Rejected",
  //       CurrentApproverId: null,
  //       ApproverComment: vendor.Comments || ""
  //     }, props);

  //     alert("❌ Rejected Successfully");
  //     history.push("/ApprovalDashboard");

  //   } catch (err) {
  //     console.error(err);
  //   }
  // };

  const handleReject = async () => {

    // ✅ Comments validation
    if (!vendor.Comments || vendor.Comments.trim() === "") {
      alert("Comments are required for rejection.");
      return;
    }

    const sp = await spCrudOps;

    try {

      // 🔥 1. GET ITEM
      const items = await sp.getData(
        "EMDDetails",
        "Id,ApprovalMatrix,WFHistory,CommentHistory",
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
      let matrix =
        typeof item.ApprovalMatrix === "string"
          ? JSON.parse(item.ApprovalMatrix)
          : item.ApprovalMatrix || [];

      const currentIndex = matrix.findIndex((x: { Status: string }) => x.Status === "Pending");

      if (currentIndex !== -1) {
        matrix[currentIndex].Status = "Rejected";
      }

      // ===============================
      // 🔥 WF HISTORY + COMMENT HISTORY
      // ===============================

      let prevHistory: any[] = [];
      let prevComments: any[] = [];

      try {
        prevHistory =
          typeof item.WFHistory === "string"
            ? JSON.parse(item.WFHistory)
            : item.WFHistory || [];

        prevComments =
          typeof item.CommentHistory === "string"
            ? JSON.parse(item.CommentHistory)
            : item.CommentHistory || [];

      } catch {
        prevHistory = [];
        prevComments = [];
      }

      // ✅ Current user
      const currentUser =
        props.context.pageContext.user.displayName ||
        props.context.pageContext.user.email;

      // ✅ Date format
      const today = new Date();
      const formattedDate =
        String(today.getDate()).padStart(2, '0') + '/' +
        String(today.getMonth() + 1).padStart(2, '0') + '/' +
        today.getFullYear();

      // ✅ New history entry
      const historyEntry = {
        CurrentApprover: currentUser,
        ActionTaken: "Rejected",
        Comment: vendor.Comments || "",
        Date: formattedDate,
        CurrentStatus: "Rejected"
      };

      // 🔥 PUSH
      prevHistory.push(historyEntry);
      prevComments.push(historyEntry);

      // 🔥 FINAL UPDATE
      await sp.updateData("EMDDetails", itemId, {
        ApprovalMatrix: JSON.stringify(matrix),
        Status: "Rejected",
        PendingAt: "Rejected",
        CurrentApproverId: null,
        ApproverComment: vendor.Comments || "",

        // ✅ HISTORY SAVE
        WFHistory: JSON.stringify(prevHistory),
        CommentHistory: JSON.stringify(prevComments)

      }, props);

      alert("❌ Rejected Successfully");
      history.push("/ApprovalDashboard");

    } catch (err) {
      console.error(err);
      alert("Error during rejection");
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
              <h1>Approval Form </h1>
            </div>
            {(title || status) ? (
              <Section title="Request Summary">
                <Grid>
                  <Field label="EMD Request No.">
                    <input type="text" value={title} readOnly />
                  </Field>
                  <Field label="Current Status">
                    <input type="text" value={status || "Pending for Approval"} readOnly />
                  </Field>
                </Grid>
              </Section>
            ) : null}
            <div className="emd-hierarchy">

              {/* Initiator */}
              <div className="emd-step active-step">
                {employee.EmployeeName}
              </div>

              {approverMatrix.map((step, index) => {

                let stepClass = "emd-step";

                const firstPending = approverMatrix.findIndex(s => s.Status === "Pending");
                const isLast = index === approverMatrix.length - 1;

                if (step.Status === "Approved") {
                  stepClass += " green";
                }
                else if (step.Status === "Rejected") {
                  stepClass += " red";
                }
                else if (index === firstPending) {
                  stepClass += " orange"; // current
                }

                return (
                  <div key={index} className={stepClass}>
                    <div>{step.Approver || step.ApproverName}</div>
                  </div>
                );
              })}

            </div>
            <div className='borderedbox'>
              {/* 🔹 Section Title */}
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
                    <Dropdown
                      className='formtext-control'
                      defaultValue={"All Selected"}
                      options={vendorCodeOptions}
                      selectedKey={vendorCodeKey}
                      disabled
                      onChange={(e, option) => {
                        if (option) {
                          setVendorCodeKey(option.key as number);
                          setVendorCode(option.text);
                        }
                      }}
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>Vendor Name </label>
                    <Dropdown
                      className='formtext-control'
                      options={vendorNameOptions}
                      selectedKey={vendorName}
                      onChange={(e, option) => option && setVendorName(option.key as number)}
                      disabled
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>Vendor Site </label>
                    <Dropdown
                      className='formtext-control'
                      options={vendorSiteOptions}
                      selectedKey={vendorSiteKey}
                      onChange={(e, option) => {
                        if (option) {
                          setVendorSiteKey(option.key as number);
                          setVendorSite(option.text);
                        }
                      }}
                      disabled
                    />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Contract Type </label>
                    <input
                      type="text" className='form-control'
                      value={
                        contractTypeOptions.find(o => o.key === contractType)?.text || ""
                      }
                      readOnly
                    />

                  </div>
                  <div className='col-md-4'>
                    <label className="font">Tender No </label>
                    <input value={vendor.TenderNo} className='form-control' readOnly />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Tender Date </label>
                    <input type="date" value={vendor.TenderDate} readOnly className='form-control' />

                  </div>
                </div>
                <div className='row mb-20'>
                  {/* <div className='col-md-4'>
                    <label className="font">Tender Type </label>
                    <input value={vendor.TenderAmount} readOnly className="form-control" />
                  </div> */}
                  <div className='col-md-4'>
                    <label className="font">Tender Amount </label>
                    <input value={vendor.TenderAmount} readOnly className="form-control" />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">EMD Amount </label>
                    <input value={vendor.EMDAmount} readOnly className="form-control" />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Currency </label>
                    <Dropdown
                      className='formtext-control'
                      options={currencyOptions}
                      selectedKey={currency}
                      onChange={(e, option) => option && setCurrency(option.key as number)}
                      disabled
                    />
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">Tender Closing Date </label>
                      <input type="date" value={vendor.TenderClosingDate} readOnly className="form-control" />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">EMD Percentage </label>
                      <input value={vendor.EMDPercentage} readOnly className="form-control" />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Mode of Payment </label>
                      <Dropdown
                        className='formtext-control'
                        options={modeOfPaymentOptions}
                        selectedKey={modeOfPayment}
                        onChange={(e, option) => option && setModeOfPayment(option.key as number)}
                        disabled
                      />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">Product Type </label>
                      <Dropdown
                        className='formtext-control'
                        options={productTypeOptions}
                        selectedKey={productType}
                        onChange={(e, option) => option && setProductType(option.key as number)}
                        disabled
                      />
                    </div>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Uploaded Document</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-3'>
                    <label className='font'> Existing Attachments  </label>
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
                          {wfHistory.map((item, index) => {
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

                            const getStatusClass = () => {
                              if (item.ActionTaken?.toLowerCase().includes("approved")) return "approvedRow";
                              if (item.ActionTaken?.toLowerCase().includes("rejected")) return "rejectedRow";
                              if (item.ActionTaken?.toLowerCase().includes("submitted")) return "submittedRow";
                              if (item.ActionTaken?.toLowerCase().includes("send back")) return "sendBackRow";
                              return "";
                            };

                            return (
                              <tr key={index} className={getStatusClass()}>

                                <td className="px-4 py-2">
                                  <div className="wfUser">
                                    <div className="wfAvatar">
                                      {item.CurrentApprover?.charAt(0)}
                                    </div>
                                    {item.CurrentApprover}
                                  </div>
                                </td>

                                <td className="wfAction">
                                  {item.ActionTaken}
                                </td>

                                <td className="px-4 py-2">{formatDate(item.Date)}</td>

                                <td className="px-4 py-2">{item.Comment || "-"}</td>

                              </tr>
                            );
                          })}
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
                  <div className='col-md-3'>
                    <label className='font'> Comments  </label>
                    <textarea
                      placeholder="Enter your comments here..."
                      rows={4} className="form-control"
                      onChange={(e) => setVendor(prev => ({ ...prev, Comments: e.target.value }))}
                    />
                  </div>
                </div>
              </div>
              <div className='row my-3'>
                <div className='col-md-12'>
                  <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                    <button onClick={handleApprove} disabled={actionLoading} className="submit-btn">
                      {actionLoading ? "Approving..." : "Approve"}
                    </button>
                    <button className="Rework-btn" onClick={handleReject} disabled={actionLoading}>
                      {actionLoading ? "Updating..." : "Reject"}
                    </button>
                    <a className="reset-btn" onClick={() => history.push("/ApprovalDashboard")}>
                      Exit
                    </a>
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

export default MANACApprovalForm;
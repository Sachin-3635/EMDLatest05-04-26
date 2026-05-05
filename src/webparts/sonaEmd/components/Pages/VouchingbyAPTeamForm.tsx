import React, { useEffect, useMemo, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory, useLocation } from "react-router-dom";
import { Dropdown, IDropdownOption } from "@fluentui/react";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { sp } from "@pnp/sp";

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

const VouchingbyAPTeamForm = (props: ISonaEmdProps) => {
  const history = useHistory();
  const location = useLocation();
  const spCrudOps = SPCRUDOPS();

  // ---------- Setup PnP for attachments ----------
  useEffect(() => {
    try { sp.setup({ spfxContext: (props as any).context }); } catch { }
  }, []);

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
    console.log("[Vouching] Parsed ItemId:", idStr, "=>", parsed);
    return parsed;
  }, [location.search, location.hash]);

  // ---------- Options ----------
  const [contractTypeOptions, setContractTypeOptions] = useState<IDropdownOption[]>([]);
  const [modeOfPaymentOptions, setModeOfPaymentOptions] = useState<IDropdownOption[]>([]);
  const [vendorNameOptions, setVendorNameOptions] = useState<IDropdownOption[]>([]);
  const [vendorCodeOptions, setVendorCodeOptions] = useState<IDropdownOption[]>([]);
  const [vendorSiteOptions, setVendorSiteOptions] = useState<IDropdownOption[]>([]);
  const [productTypeOptions, setProductTypeOptions] = useState<IDropdownOption[]>([]);
  const [currencyOptions, setCurrencyOptions] = useState<IDropdownOption[]>([]);
  const [tenderTypeOptions, setTenderTypeOptions] = useState<IDropdownOption[]>([]);

  // ---------- Selected keys ----------
  const [contractType, setContractType] = useState<number | undefined>();
  const [modeOfPayment, setModeOfPayment] = useState<number | undefined>();
  const [vendorName, setVendorName] = useState<number | undefined>();
  const [vendorCodeKey, setVendorCodeKey] = useState<number | undefined>();
  const [vendorSiteKey, setVendorSiteKey] = useState<number | undefined>();
  const [productType, setProductType] = useState<number | undefined>();
  const [currency, setCurrency] = useState<number | undefined>();
  const [tenderType, setTenderType] = useState<number | undefined>();

  // ---------- Misc / text ----------
  const [vendorCode, setVendorCode] = useState<string>("");
  const [vendorSite, setVendorSite] = useState<string>("");

  const [title, setTitle] = useState<string>("");
  const [currentStatus, setCurrentStatus] = useState<string>("");

  const [employee, setEmployee] = useState({
    EmployeeCode: "", EmployeeName: "", Division: "", Department: "",
    Location: "", ReportingManager: "", RM: "", HOD: "", ContactNo: "",
    EmployeeStatus: "", EmployeeEmail: "", ReportingManagerId: 0, HODId: 0,
  });

  const [vendor, setVendor] = useState({
    TenderNo: "", TenderDate: "", TenderAmount: "", EMDAmount: "",
    TenderClosingDate: "", EMDPercentage: "",
  });

  // History
  const [approverComment, setApproverComment] = useState<string>("");
  const [approverActionDate, setApproverActionDate] = useState<string>("");

  // Attachments
  const [attachments, setAttachments] = useState<Array<{ FileName: string; ServerRelativeUrl: string }>>([]);

  // AP Action
  const [apVouchingDate, setVouchingDate] = useState<string>("");
  const [apGLCode, setGLCode] = useState<string>("");
  const [apVoucherNo, setVoucherNo] = useState<string>("");
  const [apComment, setApTeamComment] = useState<string>("");

  // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&
  const [isTenderDuplicate, setIsTenderDuplicate] = useState<boolean>(false);
  // &&&&&&&&&&&&& To check Duplicate Tender Number &&&&&&&&&&&&&

  const [files, setFiles] = useState<File[]>([]);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  // Vendor 
  const [vendorAll, setVendorAll] = useState<any[]>([]);
  const [wfHistory, setWfHistory] = useState<any[]>([]);
  useEffect(() => {
    if (vendorCode && vendorSite && vendorCodeOptions.length > 0 && vendorSiteOptions.length > 0) {
      resolveVendorKeysByText(vendorCode, vendorSite);
    }
  }, [vendorCode, vendorSite, vendorCodeOptions, vendorSiteOptions]);

  // ---------- Load masters (guarded; do not block details on error) ----------
  useEffect(() => {
    (async () => {
      try {
        const spx = await spCrudOps;
        // Currency
        try {
          const data = await spx.getData("CurrencyMaster", "Title,Id,Currency,Status", "",
            "Status eq 'Active'", { column: "Id", isAscending: true }, 5000, props);
          setCurrencyOptions(data.map((m: any) => ({ key: m.Id, text: m.Currency })));
        } catch (e) { console.warn("[Vouching] Currency master error", e); }

        // Vendor
        try {
          const res = await spx.getData("VendorMaster", "ID,VendorCode,Name,VendorSite", "",
            "", { column: "ID", isAscending: true }, 5000, props);
          setVendorAll(res);
          setVendorNameOptions(res.map((v: any) => ({ key: v.ID, text: v.Name })));
          setVendorCodeOptions(res.map((v: any) => ({ key: v.ID, text: v.VendorCode })));
          setVendorSiteOptions(res.map((v: any) => ({ key: v.ID, text: v.VendorSite })));
        } catch (e) { console.warn("[Vouching] Vendor master error", e); }

        // Mode of Payment
        try {
          const res = await spx.getData("ModeofPaymentMaster", "Title,Id,Mode,Status", "",
            "Status eq 'Active'", { column: "Title", isAscending: true }, 5000, props);
          setModeOfPaymentOptions(res.map((m: any) => ({ key: m.Id, text: m.Mode })));
        } catch (e) { console.warn("[Vouching] ModeOfPayment master error", e); }

        // Product Type
        try {
          const res = await spx.getData("ProductTypeMaster", "Title,Id,ProductType,Status", "",
            "Status eq 'Active'", { column: "Id", isAscending: true }, 5000, props);
          setProductTypeOptions(res.map((m: any) => ({ key: m.Id, text: m.ProductType })));
        } catch (e) { console.warn("[Vouching] ProductType master error", e); }

        // Contract Type
        try {
          const res = await spx.getData("ContractTypeMaster", "Title,Id,ContractType,Status", "",
            "Status eq 'Active'", { column: "Id", isAscending: true }, 5000, props);
          setContractTypeOptions(res.map((c: any) => ({ key: c.Id, text: c.ContractType })));
        } catch (e) { console.warn("[Vouching] ContractType master error", e); }

        // Tender Type
        try {
          const res = await spx.getData("TenderType", "Title,Id,TenderType,Status", "",
            "Status eq 'Active'", { column: "Id", isAscending: true }, 5000, props);
          setTenderTypeOptions(res.map((t: any) => ({ key: t.Id, text: t.TenderType || t.Title })));
        } catch (e) { console.warn("[Vouching] TenderType master error", e); }

      } catch (e) {
        console.error("[Vouching] Masters load fatal", e);
      }
    })();
  }, []); // masters

  // ---------- Load details (robust) ----------
  useEffect(() => {
    if (!itemId || Number.isNaN(itemId)) {
      console.warn("[Vouching] Invalid ItemId", itemId);
      return;
    }

    (async () => {
      const details = await fetchEMDDetailsRobust(itemId);
      if (!details) {
        alert("No data found for this ItemId.");
        return;
      }

      console.log("[Vouching] Details:", details);
      setTitle(details.Title || "");
      setCurrentStatus(details.Status || "");

      // Requestor
      setEmployee({
        EmployeeCode: details.EmployeeName ? (details.EmployeeCode || "") : (details.EmployeeCode || ""),
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

      // Vendor/Tender
      setVendor({
        TenderNo: details.TenderNo || "",
        TenderDate: details.TenderDate ? String(details.TenderDate).split("T")[0] : "",
        TenderAmount: details.TenderAmount || "",
        EMDAmount: details.EMDAmount || "",
        TenderClosingDate: details.TenderClosingDate ? String(details.TenderClosingDate).split("T")[0] : "",
        EMDPercentage: details.EMDPercentage || "",
      });

      // Lookups
      setVendorName(details.VendorNameId || undefined);
      setContractType(details.ContractTypeId || undefined);
      setModeOfPayment(details.ModeofPaymentId || undefined);
      setProductType(details.ProductTypeId || undefined);
      setCurrency(details.CurrencyId || undefined);
      setTenderType(details.TenderTypeId || undefined);

      // Code/site + resolve keys
      setVendorCode(details.VendorCode || "");
      setVendorSite(details.VendorSite || "");
      resolveVendorKeysByText(details.VendorCode || "", details.VendorSite || "");

      // Approver history
      setApproverComment(details.ApproverComment || "");
      const d = details.ApproverActionDate || details.Modified;
      if (d) setApproverActionDate(String(d).split("T")[0]);


      await loadAttachments(itemId);

      // 🔥 LOAD APPROVAL MATRIX
      try {
        let matrix: any[] = [];

        if (details.ApprovalMatrix) {
          matrix =
            typeof details.ApprovalMatrix === "string"
              ? JSON.parse(details.ApprovalMatrix)
              : details.ApprovalMatrix;
        }

        if (Array.isArray(matrix)) {
          matrix.sort((a: any, b: any) => a.Seq - b.Seq);
          setApprovalMatrix(matrix);
        }

      } catch (e) {
        console.error("ApprovalMatrix parse error", e);
      }
      // 🔥 WF HISTORY LOAD
      try {
        let historyArr: any[] = [];

        if (details.WFHistory) {
          historyArr =
            typeof details.WFHistory === "string"
              ? JSON.parse(details.WFHistory)
              : details.WFHistory;
        }

        setWfHistory(Array.isArray(historyArr) ? historyArr : []);
      } catch (e) {
        console.error("WFHistory parse error", e);
        setWfHistory([]);
      }
    })();
  }, [itemId]); // only depends on itemId

  // ---------- Robust details fetch: BAL(ID)->BAL(Id)->PnP fallback ----------
  //   const fetchEMDDetailsRobust = async (id: number) => {
  //     const spx = await spCrudOps;
  //     const select =
  //       "ID,Title,Status,ApproverComment,ApproverActionDate,Modified," +
  //       "EmployeeName,EmployeeCode,Division,Location,RM,HOD,ContactNo,EmployeeStatus,Email," +
  //       "VendorCode,VendorSite," +
  //       "TenderTypeId,TenderType/TenderType," +
  //       "VendorNameId,VendorName/Title," +
  //       "ContractTypeId,ContractType/Title," +
  //       "ModeofPaymentId,ModeofPayment/Title," +
  //       "ProductTypeId,ProductType/Title," +
  //       "TenderNo,TenderDate,TenderAmount,EMDAmount,CurrencyId,Currency/Currency,TenderClosingDate,EMDPercentage," +
  //       "APVouchingDate,APGLCode,APVoucherNo,APComment";

  //     const expand = "VendorName,ContractType,ProductType,Currency,ModeofPayment,TenderType";

  //     const tries = [
  //       { filter: `ID eq ${id}`, order: { column: "ID", isAscending: true } },
  //       { filter: `Id eq ${id}`, order: { column: "Id", isAscending: true } },
  //     ];

  //     for (const t of tries) {
  //       try {
  //         console.log("[Vouching] getData with", t.filter);
  //         const res = await spx.getData("EMDDetails", select, expand, t.filter, t.order, 1, props);
  //         if (res && res.length) return res[0];
  //       } catch (e) {
  //         console.warn("[Vouching] getData error with", t.filter, e);
  //       }
  //     }

  //     // PnP fallback (bypass BAL completely)
  //     try {
  //       console.log("[Vouching] Fallback PnP getById");
  //       const pnp = await sp.web.lists.getByTitle("EMDDetails").items.getById(id)
  //         .select(select) // select with expands is not native-friendly; split:
  //         .expand("VendorName,ContractType,ProductType,Currency,ModeofPayment,TenderType")
  //         .get();
  //       return pnp;
  //     } catch (e) {
  //       console.error("[Vouching] PnP fallback error", e);
  //       return null;
  //     }
  //   };

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


  // ---------- Robust details fetch: BAL(ID)->BAL(Id)->PnP fallback ----------
  const fetchEMDDetailsRobust = async (id: number) => {
    const spx = await spCrudOps;

    // ⛔ Removed ApproverActionDate from select
    const select =
      "ID,Title,Status,ApproverComment,Modified," +
      "EmployeeName,EmployeeCode,Division,Department,Location,RM,HOD,ContactNo,EmployeeStatus,Email," +
      "VendorCode,VendorSite," +
      "TenderTypeId,TenderType/TenderType," +
      "VendorNameId,VendorName/Title," +
      "ContractTypeId,ContractType/Title," +
      "ModeofPaymentId,ModeofPayment/Title," +
      "ProductTypeId,ProductType/Title," +
      "TenderNo,TenderDate,TenderAmount,EMDAmount,CurrencyId,Currency/Currency,TenderClosingDate,EMDPercentage," +
      "VouchingDate,GLCode,VoucherNo,APTeamComment,ApprovalMatrix,WFHistory";

    const expand = "VendorName,ContractType,ProductType,Currency,ModeofPayment,TenderType";

    const tries = [
      { filter: `ID eq ${id}`, order: { column: "ID", isAscending: true } },
      { filter: `Id eq ${id}`, order: { column: "Id", isAscending: true } },
    ];

    for (const t of tries) {
      try {
        console.log("[Vouching] getData with", t.filter);
        const res = await spx.getData("EMDDetails", select, expand, t.filter, t.order, 1, props);
        if (res && res.length) return res[0];
      } catch (e) {
        console.warn("[Vouching] getData error with", t.filter, e);
      }
    }

    // PnP fallback ( also without ApproverActionDate)
    try {
      console.log("[Vouching] Fallback PnP getById");
      const pnp = await sp.web.lists
        .getByTitle("EMDDetails")
        .items.getById(id)
        .select(
          "ID", "Title", "Status", "ApproverComment", "Modified",
          "EmployeeName", "EmployeeCode", "Department", "Division", "Location", "RM", "HOD", "ContactNo", "EmployeeStatus", "Email",
          "VendorCode", "VendorSite",
          "TenderTypeId", "VendorNameId", "ContractTypeId", "ModeofPaymentId", "ProductTypeId",
          "TenderNo", "TenderDate", "TenderAmount", "EMDAmount", "CurrencyId", "TenderClosingDate", "EMDPercentage",
          "VouchingDate", "GLCode", "VoucherNo", "APTeamComment",
          "Currency/Currency",
          "VendorName/Title",
          "ContractType/Title",
          "ProductType/Title",
          "ModeofPayment/Title",
          "TenderType/TenderType",
          "WFHistory"
        )
        .expand("VendorName", "ContractType", "ProductType", "Currency", "ModeofPayment", "TenderType")
        .get();

      return pnp;
    } catch (e) {
      console.error("[Vouching] PnP fallback error", e);
      return null;
    }
  };


  // ---------- Resolve vendor keys from text ----------
  const resolveVendorKeysByText = (codeText: string, siteText: string) => {
    if (vendorCodeOptions.length) {
      const codeOpt = vendorCodeOptions.find((o) => o.text === codeText);
      if (codeOpt) setVendorCodeKey(codeOpt.key as number);
    }
    if (vendorSiteOptions.length) {
      const siteOpt = vendorSiteOptions.find((o) => o.text === siteText);
      if (siteOpt) setVendorSiteKey(siteOpt.key as number);
    }
  };

  // ---------- Load attachments ----------
  const loadAttachments = async (id: number) => {
    try {
      const files = await sp.web.lists.getByTitle("EMDDetails").items.getById(id).attachmentFiles();
      console.log("[Vouching] Attachments:", files);
      setAttachments(files || []);
    } catch (e) {
      console.warn("[Vouching] Attachments fetch failed", e);
      setAttachments([]);
    }
  };

  //--------Update Matrix----------
  const getUpdatedApprovalMatrix = async (oldMatrix: any[]) => {

    const sp = await spCrudOps;

    const master = await sp.getData(
      "EMDApprovalMatrix",
      "ID,Role/RoleName,Approver/ID,Approver/Title",
      "Role,Approver",
      "RequestType eq 'EMD Approval'",
      { column: "ID", isAscending: true },
      500,
      props
    );

    const masterMatrix = master.map((x: any) => ({
      Role: x.Role?.RoleName,
      ApproverID: String(x.Approver?.ID),
      Approver: x.Approver?.Title,
    }));

    return oldMatrix.map((oldItem: any, index: number) => {

      // ✅ KEEP RM & HOD
      if (index === 0 || index === 1) {
        return oldItem;
      }

      const masterItem = masterMatrix.find((m: { Role: any; }) => m.Role === oldItem.Role);

      if (!masterItem) return oldItem;

      return {
        ...oldItem,
        ApproverID: masterItem.ApproverID,
        Approver: masterItem.Approver,
        Status: oldItem.Status
      };
    });
  };

  // ---------- Submit ----------
  const onsubmit = async () => {
    try {
      if (!itemId || Number.isNaN(itemId)) return alert("Invalid item id.");
      if (!apVouchingDate) return alert("Vouching Date is required.");
      if (!apGLCode.trim()) return alert("GL Code is required.");
      if (!apVoucherNo.trim()) return alert("Voucher No. is required.");
      if (!apComment.trim()) return alert("Comments are required.");

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

      let matrix: any[] = [];

      try {
        matrix =
          typeof items[0].ApprovalMatrix === "string"
            ? JSON.parse(items[0].ApprovalMatrix)
            : items[0].ApprovalMatrix || [];
      } catch {
        matrix = [];
      }

      // 🔥 2. UPDATE MATRIX FROM MASTER
      matrix = await getUpdatedApprovalMatrix(matrix);

      // 🔥 3. FIND CURRENT
      // 🔥 3. FIND CURRENT
      const currentIndex = matrix.findIndex(x => x.Status === "Pending");

      if (currentIndex === -1) {
        alert("No pending approver");
        return;
      }

      // 🔥 4. APPROVE CURRENT
      matrix[currentIndex] = {
        ...matrix[currentIndex],
        Status: "Approved"
      };

      const nextIndex = currentIndex + 1;

      let payload: any = {
        VouchingDate: apVouchingDate,
        GLCode: apGLCode.trim(),
        VoucherNo: apVoucherNo.trim(),
        APTeamComment: apComment.trim()
      };

      // 🔥 5. NEXT APPROVER
      if (nextIndex < matrix.length) {

        const next = matrix[nextIndex];

        // ✅ VERY IMPORTANT
        matrix[nextIndex] = {
          ...next,
          Status: "Pending"
        };

        payload.CurrentApproverId = Number(next.ApproverID);
        payload.Status = "Pending for Payment";
        payload.PendingAt = `Pending at ${next.Approver}`;

      } else {

        payload.CurrentApproverId = null;
        payload.Status = "Completed";
        payload.PendingAt = "Completed";
      }

      // 🔥 6. SAVE MATRIX
      payload.ApprovalMatrix = JSON.stringify(matrix);
      // 🔥 ADD WF HISTORY ENTRY
      // 🔥 WF HISTORY UPDATE
      let prevHistory: any[] = [];

      try {
        prevHistory =
          typeof items[0].WFHistory === "string"
            ? JSON.parse(items[0].WFHistory)
            : items[0].WFHistory || [];
      } catch {
        prevHistory = [];
      }

      // ✅ Current user
      const currentUser =
        props.context?.pageContext?.user?.displayName || "AP Team";

      const today = new Date();

      const formattedDate =
        String(today.getDate()).padStart(2, '0') + '/' +
        String(today.getMonth() + 1).padStart(2, '0') + '/' +
        today.getFullYear();

      // ✅ New entry
      const newEntry = {
        CurrentApprover: currentUser,
        ActionTaken: "AP Vouching Completed",
        Comment: apComment,
        Date: formattedDate
      };

      // ✅ Push new history
      prevHistory.push(newEntry);

      // ✅ Save back
      payload.WFHistory = JSON.stringify(prevHistory);
      // 🔥 6. UPDATE
      await spx.updateData("EMDDetails", itemId, payload, props);

      // if (files.length > 0) {
      //   await Promise.all(
      //     files.map((file) => spx.addAttchmentInList(file, "EMDDetails", itemId, file.name, props))
      //   );
      // }

      if (files.length > 0) {
        for (const file of files) {

          const name = file.name.substring(0, file.name.lastIndexOf(".")) || file.name;
          const ext = file.name.substring(file.name.lastIndexOf("."));

          const now = new Date();

          const time =
            ("0" + now.getHours()).slice(-2) +
            ("0" + now.getMinutes()).slice(-2) +
            ("0" + now.getSeconds()).slice(-2);

          const uniqueName = `${name}_${time}${ext}`;

          await spx.addAttchmentInList(
            file,
            "EMDDetails",
            itemId,
            uniqueName,
            props
          );
        }
      }

      alert("✅ Vouching saved and moved to Pending for Payment.");
      history.goBack();
      // history.push("/TreasuryLandingPage");
    } catch (e) {
      console.error("[Vouching] Submit error", e);
      alert("Something went wrong. Please check console.");
    }
  };

  return (
    <div className="forex-wrapper">
      {/* Small debug strip — remove after verifying */}
      {/* <div style={{ background: "#fff8dc", padding: "6px 10px", border: "1px dashed #e0c26d", marginBottom: 8 }}>
        <b>DEBUG</b>: ItemId = {String(itemId)} | Title = {title || "-"}
      </div> */}

      {/* ================= HEADER ================= */}
      <div className="forex-header"><h2>Vouching by AP Team</h2></div>

      <div className="forex-card">
        {/* 1) Summary */}
        <Section title="Request Summary">
          <Grid>
            <Field label="EMD Request No."><input type="text" value={title} readOnly /></Field>
            <Field label="Current Status"><input type="text" value={currentStatus || "-"} readOnly /></Field>
          </Grid>
        </Section>

        {/* ================= APPROVAL HIERARCHY ================= */}
        <div className="headerApproval">
          <div className="approvalFlow">

            {/* Initiator */}
            <div className="flowStep green">
              {employee.EmployeeName || "Initiator"}
            </div>

            {approvalMatrix.map((step, index) => {

              let stepClass = "grey";

              const firstPending = approvalMatrix.findIndex(
                (s) => s.Status === "Pending"
              );

              // 🔴 Rejected
              if (step.Status === "Rejected") {
                stepClass = "red";
              }

              // 🟢 Approved
              else if (step.Status === "Approved") {
                stepClass = "green";
              }

              // 🟠 Current Pending
              else if (index === firstPending) {
                stepClass = "orange";
              }

              return (
                <div key={index} className={`flowStep ${stepClass}`}>
                  {step.ApproverName}
                </div>
              );
            })}

          </div>
        </div>

        {/* 2) Requestor Information */}
        <Section title="Requestor Information">
          <Grid style={{ marginTop: "20px" }}>
            <Field label="Employee Code"><input type="text" value={employee.EmployeeCode} readOnly /></Field>
            <Field label="Employee Name"><input type="text" value={employee.EmployeeName} readOnly /></Field>
            <Field label="Division"><input type="text" value={employee.Division} readOnly /></Field>
            <Field label="Location"><input type="text" value={employee.Location} readOnly /></Field>
            <Field label="Reporting Manager"><input type="text" value={employee.RM || employee.ReportingManager} readOnly /></Field>
            <Field label="HOD"><input type="text" value={employee.HOD} readOnly /></Field>
            <Field label="Contact No"><input type="text" value={employee.ContactNo} readOnly /></Field>
            <Field label="Employee Status"><input type="text" value={employee.EmployeeStatus} readOnly /></Field>
            <Field label="Department"><input type="text" value={employee.Department} readOnly /></Field>
            <Field label="Employee Email" full><input type="email" value={employee.EmployeeEmail} readOnly /></Field>
          </Grid>
        </Section>

        {isTenderDuplicate && (
          <section>
            <h5 style={{ color: "green" }}>
              This Tender No. is available with another EMD request.
            </h5>
          </section>
        )}

        {/* 3) EMD Request Details */}
        <Section title="EMD Request Details">
          <Grid>
            <Field label="Vendor Code">
              <Dropdown options={vendorCodeOptions} selectedKey={vendorCodeKey} disabled />
            </Field>
            <Field label="Vendor Name">
              <Dropdown options={vendorNameOptions} selectedKey={vendorName} disabled />
            </Field>
            <Field label="Vendor Site">
              <Dropdown options={vendorSiteOptions} selectedKey={vendorSiteKey} disabled />
            </Field>
            <Field label="Contract Type">
              <Dropdown options={contractTypeOptions} selectedKey={contractType} disabled />
            </Field>
            <Field label="Tender No."><input value={vendor.TenderNo} readOnly /></Field>
            <Field label="Tender Date"><input type="date" value={vendor.TenderDate} readOnly /></Field>
            <Field label="Tender Type">
              <Dropdown options={tenderTypeOptions} selectedKey={tenderType} disabled />
            </Field>
            <Field label="Tender Amount"><input value={vendor.TenderAmount} readOnly /></Field>
            <Field label="EMD Amount"><input value={vendor.EMDAmount} readOnly /></Field>
            <Field label="Currency">
              <Dropdown options={currencyOptions} selectedKey={currency} disabled />
            </Field>
            <Field label="Tender Closing Date"><input type="date" value={vendor.TenderClosingDate} readOnly /></Field>
            <Field label="EMD Percentage"><input value={vendor.EMDPercentage} readOnly /></Field>
            <Field label="Mode Of Payment">
              <Dropdown options={modeOfPaymentOptions} selectedKey={modeOfPayment} disabled />
            </Field>
            <Field label="Product Type">
              <Dropdown options={productTypeOptions} selectedKey={productType} disabled />
            </Field>
          </Grid>
        </Section>
        <Section title="Workflow History">

          <div className="wfTableWrapper">

            <table className="wfTable">
              <thead>
                <tr>
                  <th>Action By</th>
                  <th>Action Taken</th>
                  <th>Date</th>
                  <th>Comment</th>
                </tr>
              </thead>

              <tbody>
                {wfHistory.length > 0 ? (
                  wfHistory.map((item, index) => {

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
        {/* 4) AP Action */}
        <Section title="Action">
          <Grid>
            <Field
              // label="Vouching Date"
              label={
                <span>Vouching Date<span style={{ color: "red" }}>*</span></span>
              }

            >
              <input type="date" value={apVouchingDate} onChange={(e) => setVouchingDate(e.target.value)} />
            </Field>
            <Field
              //  label="GL Code*"
              label={
                <span>GL Code<span style={{ color: "red" }}>*</span></span>
              }
            >
              <input value={apGLCode} onChange={(e) => setGLCode(e.target.value)} placeholder="Enter GL Code" />
            </Field>
            <Field
              // label="Voucher No.*"
              label={
                <span>Voucher No.<span style={{ color: "red" }}>*</span></span>
              }
            >
              <input value={apVoucherNo} onChange={(e) => setVoucherNo(e.target.value)} placeholder="Enter Voucher No." />
            </Field>
            <Field
              // label="Comments*"
              label={
                <span>Comments<span style={{ color: "red" }}>*</span></span>
              }
            >
              <textarea rows={3} value={apComment} onChange={(e) => setApTeamComment(e.target.value)} placeholder="Enter your comments" />
            </Field>
          </Grid>
        </Section>

        {/* 5) Upload (AP supporting) */}
        <Section title="Upload Documents">
          <Grid>
            <Field label="Attach">
              <input type="file" multiple onChange={(e) => { if (e.target.files) setFiles(Array.from(e.target.files)); }} />
            </Field>
          </Grid>
        </Section>



        {/* 6) Uploaded Documents (BOTTOM) */}
        <Section title="Uploaded Documents">
          <Grid>
            <Field label="Attach">
              {attachments.length === 0 ? <div>-</div> : (
                <ul style={{ margin: 0, paddingLeft: 18 }}>
                  {attachments.map(a => (
                    <li key={a.ServerRelativeUrl}>
                      <a href={a.ServerRelativeUrl} target="_blank" rel="noreferrer">{a.FileName}</a>
                    </li>
                  ))}
                </ul>
              )}
            </Field>
          </Grid>
        </Section>

        {/* 7) Workflow History (BOTTOM) */}
        {/* <Section title="Workflow History">
          <Grid>
            <Field label="Approval By"><input type="text" value="MANAC Team" readOnly /></Field>
            <Field label="Action Taken"><input type="text" value="Approved" readOnly /></Field>
            <Field label="Action Date"><input type="text" value={approverActionDate || "-"} readOnly /></Field>
            <Field label="Supporting Docs" full>
              {attachments.length === 0 ? <div>-</div> : (
                <ul style={{ margin: 0, paddingLeft: 18 }}>
                  {attachments.map(a => (
                    <li key={`wf-${a.ServerRelativeUrl}`}>
                      <a href={a.ServerRelativeUrl} target="_blank" rel="noreferrer">{a.FileName}</a>
                    </li>
                  ))}
                </ul>
              )}
            </Field>
          </Grid>
        </Section> */}


        {/* 8) Comment History (BOTTOM) */}
        {/* <Section title="Comment History">
          <Grid>
            <Field label="Comment By"><input type="text" value="MANAC Team" readOnly /></Field>
            <Field label="Comment"><textarea value={approverComment || "-"} readOnly rows={3} /></Field>
            <Field label="Comment Date"><input type="text" value={approverActionDate || "-"} readOnly /></Field>
          </Grid>
        </Section> */}

        <div className="button-row">
          <button className="btn-submit" onClick={onsubmit}>Submit</button>
          <button className="btn-exit" onClick={() => history.goBack()}>Exit</button>
        </div>
      </div>
    </div>
  );
};

export default VouchingbyAPTeamForm;
import React, { useEffect, useMemo, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory, useLocation } from 'react-router-dom';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import SPCRUDOPS from "../../service/BAL/spcrud";
import logo from "../../assets/sona-comstarlogo.png";

// Simple UI helpers
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

const EMDRequestFormEdit = (props: ISonaEmdProps) => {
  const history = useHistory();
  const location = useLocation();
  const spCrudOps = SPCRUDOPS();

  // ---------- Parse ItemId from hash query (#/route?ItemId=12) ----------
  // const itemId = useMemo(() => {
  //   // location.hash looks like: "#/EMDRequestFormView?ItemId=12"
  //   const hash = location.hash || window.location.hash || "";
  //   const qIndex = hash.indexOf("?");
  //   const search = qIndex >= 0 ? hash.substring(qIndex) : "";
  //   const qs = new URLSearchParams(search);
  //   const idStr = qs.get("ItemId");
  //   return idStr ? parseInt(idStr, 10) : NaN;
  // }, [location.hash]);


  const itemId = useMemo(() => {
    // React Router v5 sets location.search correctly even under HashRouter
    const qs = new URLSearchParams(location.search || window.location.search);
    const idStr = qs.get("ItemId");
    return idStr ? parseInt(idStr, 10) : NaN;
  }, [location.search]);



  // ---------- Options ----------
  const [contractTypeOptions, setContractTypeOptions] = useState<IDropdownOption[]>([]);
  const [modeOfPaymentOptions, setModeOfPaymentOptions] = useState<IDropdownOption[]>([]);
  const [vendorNameOptions, setVendorNameOptions] = useState<IDropdownOption[]>([]);
  const [vendorCodeOptions, setVendorCodeOptions] = useState<IDropdownOption[]>([]);
  const [vendorSiteOptions, setVendorSiteOptions] = useState<IDropdownOption[]>([]);
  const [productTypeOptions, setProductTypeOptions] = useState<IDropdownOption[]>([]);
  const [currencyOptions, setCurrencyOptions] = useState<IDropdownOption[]>([]);
  const [tenderTypeOptions, setTenderTypeOptions] = useState<IDropdownOption[]>([]); // NEW

  // ---------- Selected keys ----------
  const [contractType, setContractType] = useState<number | undefined>();
  const [modeOfPayment, setModeOfPayment] = useState<number | undefined>();
  const [vendorName, setVendorName] = useState<number | undefined>();
  const [vendorCodeKey, setVendorCodeKey] = useState<number | undefined>();
  const [vendorCode, setVendorCode] = useState<string>("");
  const [vendorSiteKey, setVendorSiteKey] = useState<number | undefined>();
  const [vendorSite, setVendorSite] = useState<string>("");
  const [productType, setProductType] = useState<number | undefined>();
  const [currency, setCurrency] = useState<number | undefined>();
  const [tenderType, setTenderType] = useState<number | undefined>(); // NEW
  const [duplicateWarning, setDuplicateWarning] = useState("");

  // ---------- Misc ----------
  const [files, setFiles] = useState<File[]>([]);
  const [title, setTitle] = useState<string>(""); // EMD Request No. (Title)
  const [approvalType, setApprovalType] = useState<any[]>([]);
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
  });

  // Cache VendorMaster rows to auto-fill when code changes
  const [vendorAll, setVendorAll] = useState<any[]>([]);

  // ---------- Load masters ----------
  useEffect(() => {
    getCurrencyData();
    getVendorDataMaster();
    getContractTypeData();
    getModeOfPaymentData();
    getProductTypeData();
    getTenderTypeData();
    loadApproverMatrix();
  }, []);

  // ---------- After masters ready, load the item by id ----------
  // ---------- Resolve vendor keys from text ----------
  const resolveVendorKeysByText = (codeText: string, siteText: string) => {
    if (vendorCodeOptions.length) {
      const codeOpt = vendorCodeOptions.find(o => o.text === codeText);
      if (codeOpt) setVendorCodeKey(codeOpt.key as number);
    }
    if (vendorSiteOptions.length) {
      const siteOpt = vendorSiteOptions.find(o => o.text === siteText);
      if (siteOpt) setVendorSiteKey(siteOpt.key as number);
    }
  };

  useEffect(() => {
    if (!itemId || Number.isNaN(itemId)) return;

    // For vendorCode/site key resolution we need options.
    // We can still set lookup IDs directly; selectedKey will bind when options arrive.
    fetchEMDDetails(itemId).then(details => {
      if (!details) return;

      // Title (EMD Request No.)
      setTitle(details.Title || "");

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
        HODId: details.HODId || 0
      });

      // Vendor/Tender fields
      setVendor({
        TenderNo: details.TenderNo || "",
        TenderDate: details.TenderDate ? details.TenderDate.split("T")[0] : "",
        TenderAmount: details.TenderAmount || "",
        EMDAmount: details.EMDAmount || "",
        TenderClosingDate: details.TenderClosingDate ? details.TenderClosingDate.split("T")[0] : "",
        EMDPercentage: details.EMDPercentage || ""
      });

      // Lookup dropdowns (IDs)
      setVendorName(details.VendorNameId || undefined);
      setContractType(details.ContractTypeId || undefined);
      setModeOfPayment(details.ModeofPaymentId || undefined);
      setProductType(details.ProductTypeId || undefined);
      setCurrency(details.CurrencyId || undefined);
      setTenderType(details.TenderTypeId || undefined);

      // Vendor Code/Site (text + resolve to matching key)
      setVendorCode(details.VendorCode || "");
      setVendorSite(details.VendorSite || "");

      // try to resolve keys by matching text with options (once options arrive, this runs again)
      resolveVendorKeysByText(details.VendorCode || "", details.VendorSite || "");
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [itemId, vendorCodeOptions.length, vendorSiteOptions.length]);

  // Auto-calc EMD %
  useEffect(() => {
    const tender = parseFloat(vendor.TenderAmount || "0");
    const emd = parseFloat(vendor.EMDAmount || "0");
    if (tender > 0) {
      const percent = ((emd / tender) * 100).toFixed(2);
      setVendor(prev => ({ ...prev, EMDPercentage: percent }));
    }
  }, [vendor.TenderAmount, vendor.EMDAmount]);

  // ---------- Master data loaders ----------
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

    setVendorAll(res);
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

  const getTenderTypeData = async () => {
    const sp = await spCrudOps;
    const res = await sp.getData(
      "TenderType",
      "Title,Id,TenderType,Status",
      "",
      "Status eq 'Active'",
      { column: "Id", isAscending: true },
      5000,
      props
    );
    setTenderTypeOptions(res.map((t: any) => ({ key: t.Id, text: t.TenderType || t.Title })));
  };

  // ---------- Fetch details ----------
  const fetchEMDDetails = async (id: number) => {
    const sp = await spCrudOps;
    const res = await sp.getData(
      "EMDDetails",
      "Id,Title,EmployeeName,EmployeeCode,Division,Department,Location,RM,HOD,ContactNo,EmployeeStatus,Email,VendorCode,TenderTypeId,TenderType/TenderType,VendorNameId,VendorName/Title,VendorSite,ContractTypeId,ContractType/Title,ModeofPaymentId,ModeofPayment/Title,ProductTypeId,ProductType/Title,TenderNo,TenderDate,TenderAmount,EMDAmount,CurrencyId,Currency/Currency,TenderClosingDate,EMDPercentage",
      "VendorName,ContractType,ProductType,Currency,ModeofPayment,TenderType",
      `Id eq ${id}`, // use Id here
      { column: "Id", isAscending: true },
      1,
      props
    );
    return res && res.length ? res[0] : null;
  };


  // ---------- Duplicate tender check excluding current item ----------
  const checkDuplicateTenderNumberExcludingSelf = async (tenderNo: string, id: number) => {
    if (!tenderNo) return false;
    const sp = await spCrudOps;
    const res = await sp.getData(
      "EMDDetails",
      "Id,TenderNo",
      "",
      `(TenderNo eq '${tenderNo.replace(/'/g, "''")}') and (Id ne ${id})`,
      { column: "Id", isAscending: true },
      1,
      props
    );
    return res.length > 0;
  };

  async function loadApproverMatrix() {
    try {
      const sp = await spCrudOps;

      const data = await sp.getData(
        "EMDApprovalMatrix",
        "ID,ApprovalType,Level/Level,Level/ID,Role/RoleName,Approver/ID,Approver/Title,Approver/EMail",
        "Level,Role,Approver",
        "RequestType eq 'EMD Approval'",
        { column: "Level", isAscending: true },
        5000,
        props
      );

      console.log("EMD Approval Matrix:", data);

      setApprovalType(data);

    } catch (err) {
      console.error("Error loading EMDApprovalMatrix:", err);
    }
  }

  // ---------- Submit (UPDATE existing) ----------
  const onsubmit = async () => {
    try {
      if (!itemId || Number.isNaN(itemId)) {
        alert("Invalid item id.");
        return;
      }

      const sp = await spCrudOps;

      // Validate minimal
      if (!vendorCode) {
        alert("Vendor Code is required.");
        return;
      }
      if (!vendor.TenderNo) {
        alert("Tender No. is required.");
        return;
      }

      // Duplicate tender check excluding this item
      const dup = await checkDuplicateTenderNumberExcludingSelf(vendor.TenderNo, itemId);
      if (dup) {
        alert("❌ Another request with same Tender No. already exists!");
        return;
      }

      const rm = {
        Seq: 1,
        Role: "RM",
        ApproverID: employee.ReportingManagerId,
        ApproverName: employee.ReportingManager,
        Status: "Pending"
      };

      // 🔹 HOD (from EmployeeMaster)
      const hod = {
        Seq: 2,
        Role: "HOD",
        ApproverID: employee.HODId,
        ApproverName: employee.HOD,
        Status: "Waiting"
      };

      // 🔹 Remaining approvers from EMDApprovalMatrix
      const otherApprovers = approvalType.map((x: any, index: number) => ({
        Seq: index + 3,
        Role: x.Role?.RoleName,
        ApproverID: x.Approver?.ID,
        ApproverName: x.Approver?.Title,
        Status: "Waiting"
      }));

      // 🔹 Final Approval Matrix
      const approvalMatrix = [rm, hod, ...otherApprovers];

      console.log("Final ApprovalMatrix:", approvalMatrix);

      const today = new Date();

      const formattedDate =
        String(today.getDate()).padStart(2, '0') + '/' +
        String(today.getMonth() + 1).padStart(2, '0') + '/' +
        today.getFullYear();

      const currentUser =
        props.context.pageContext.user.displayName ||
        props.context.pageContext.user.email;

      // ✅ WFHistory entry
      const initialWFHistory = [
        {
          CurrentApprover: currentUser,
          ActionTaken: "Re-Submitted",
          Comment: "Request Created",
          Date: formattedDate,
          CurrentStatus: "Pending for Approval"
        }
      ];
      // Update the existing record
      await sp.updateData(
        "EMDDetails",
        itemId,
        {
          // ----- Requestor (usually unchanged on edit, but safe to send) -----
          EmployeeName: employee.EmployeeName,
          EmployeeCode: employee.EmployeeCode,
          Division: employee.Division,
          Location: employee.Location,
          RM: employee.RM || employee.ReportingManager,
          HOD: employee.HOD,
          ContactNo: employee.ContactNo,
          EmployeeStatus: employee.EmployeeStatus,
          Department: employee.Department,
          Email: employee.EmployeeEmail,

          // ----- Vendor -----
          VendorCode: vendorCode,
          VendorNameId: vendorName ? Number(vendorName) : null, // lookup id
          VendorSite: vendorSite,

          // ----- Masters -----
          ContractTypeId: contractType ? Number(contractType) : null,
          ModeofPaymentId: modeOfPayment ? Number(modeOfPayment) : null,
          ProductTypeId: productType ? Number(productType) : null,
          CurrencyId: currency ? Number(currency) : null,
          TenderTypeId: tenderType ? Number(tenderType) : null,

          // ----- Tender -----
          TenderNo: vendor.TenderNo,
          TenderDate: vendor.TenderDate || null,
          TenderAmount: vendor.TenderAmount,
          EMDAmount: vendor.EMDAmount,
          TenderClosingDate: vendor.TenderClosingDate || null,
          EMDPercentage: vendor.EMDPercentage,
          ApprovalMatrix: JSON.stringify(approvalMatrix),
          CurrentApproverId: approvalMatrix[0]?.ApproverID,
          WFHistory: JSON.stringify(initialWFHistory),
          CommentHistory: JSON.stringify(initialWFHistory),
        },
        props
      );

      // Attachments: add any new ones
      // if (files.length > 0) {
      //   await Promise.all(
      //     files.map(file => sp.addAttchmentInList(file, "EMDDetails", itemId, file.name, props))
      //   );
      // }

      if (files && files.length > 0) {
        for (const file of files) {

          const name = file.name.substring(0, file.name.lastIndexOf(".")) || file.name;
          const ext = file.name.substring(file.name.lastIndexOf("."));

          const now = new Date();

          const time =
            ("0" + now.getHours()).slice(-2) +
            ("0" + now.getMinutes()).slice(-2) +
            ("0" + now.getSeconds()).slice(-2);

          const uniqueName = `${name}_${time}${ext}`;

          await sp.addAttchmentInList(
            file,
            "EMDDetails",
            itemId,
            uniqueName,
            props
          );
        }
      }



      alert("✅ Request updated successfully!");
      history.push("/");

    } catch (error) {
      console.error("❌ Error updating data:", error);
      alert("Something went wrong. Please check console.");
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
              <h1>Edit EMD Issuance </h1>
            </div>
            <div className="approvalFlow">

              {/* Initiator */}
              <div className="flowStep green">
                {employee.EmployeeName || "Initiator"}
              </div>

              {approverMatrix.map((step, index) => {

                let stepClass = "grey";

                const firstPending = approverMatrix.findIndex(s => s.Status === "Pending");

                // Approved
                if (step.Status === "Approved") {
                  stepClass = "green";
                }

                // Current Pending
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
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>EMD Request Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className='font'>Vendor Code </label>
                    <Dropdown
                      className='formtext-control'
                      options={vendorCodeOptions}
                      selectedKey={vendorCodeKey}
                      onChange={(e, option) => {
                        if (option) {
                          setVendorCodeKey(option.key as number);
                          setVendorCode(option.text);

                          // Auto-fill name & site from VendorMaster cache
                          const row = vendorAll.find(v => v.ID === option.key);
                          if (row) {
                            setVendorName(row.ID);
                            setVendorSiteKey(row.ID);
                            setVendorSite(row.VendorSite || "");
                          }
                        }
                      }}
                      placeholder="Select Vendor Code"
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className='font'>Vendor Name </label>
                    <Dropdown
                      className='formtext-control'
                      options={vendorNameOptions}
                      selectedKey={vendorName}
                      onChange={(e, option) => {
                        if (option) setVendorName(option.key as number);
                      }}
                      disabled
                      placeholder="Select Vendor Name"
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
                      placeholder="Select Vendor Site"
                    />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Contract Type </label>
                    <Dropdown
                      className='formtext-control'
                      options={contractTypeOptions}
                      selectedKey={contractType}
                      onChange={(e, option) => option && setContractType(option.key as number)}
                      placeholder="Select Contract Type"
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Tender No </label>
                    <input
                      value={vendor.TenderNo} className='form-control'
                      onChange={(e) => setVendor(prev => ({ ...prev, TenderNo: e.target.value }))}
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Tender Date </label>
                    <input
                      type="date" className='form-control'
                      value={vendor.TenderDate}
                      onChange={(e) => setVendor(prev => ({ ...prev, TenderDate: e.target.value }))}
                    />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Tender Type </label>
                    <Dropdown
                      className='formtext-control'
                      options={tenderTypeOptions}
                      selectedKey={tenderType}
                      onChange={(e, option) => option && setTenderType(option.key as number)}
                      placeholder="Select Tender Type"
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Tender Amount </label>
                    <input
                      value={vendor.TenderAmount} className="form-control"
                      onChange={(e) => setVendor(prev => ({ ...prev, TenderAmount: e.target.value }))}
                    />

                  </div>
                  <div className='col-md-4'>
                    <label className="font">EMD Amount </label>
                    <input
                      value={vendor.EMDAmount} className="form-control"
                      onChange={(e) => setVendor(prev => ({ ...prev, EMDAmount: e.target.value }))}
                    />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Currency </label>
                    <Dropdown
                      className='formtext-control'
                      options={currencyOptions}
                      selectedKey={currency}
                      onChange={(e, option) => option && setCurrency(option.key as number)}
                      placeholder="Select Currency"
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Tender Closing Date </label>
                    <input
                      type="date" className="form-control"
                      value={vendor.TenderClosingDate}
                      onChange={(e) => setVendor(prev => ({ ...prev, TenderClosingDate: e.target.value }))}
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">EMD Percentage </label>
                    <input
                      value={vendor.EMDPercentage} className="form-control"
                      onChange={(e) => setVendor(prev => ({ ...prev, EMDPercentage: e.target.value }))}
                      readOnly
                    />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Mode of Payment </label>
                    <Dropdown
                      className='formtext-control'
                      options={modeOfPaymentOptions}
                      selectedKey={modeOfPayment}
                      onChange={(e, option) => option && setModeOfPayment(option.key as number)}
                      placeholder="Select Mode of Payment"
                    />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Product Type </label>
                    <Dropdown
                      className='formtext-control'
                      options={productTypeOptions}
                      selectedKey={productType}
                      onChange={(e, option) => option && setProductType(option.key as number)}
                      placeholder="Select Product Type"
                    />
                  </div>

                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upload Documents</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-3'>
                    <label className='font'> Attachment  </label>
                    <input
                      type="file"
                      multiple
                      onChange={(e) => {
                        if (e.target.files) setFiles(Array.from(e.target.files));
                      }}
                    />
                  </div>
                </div>
              </div>
              <div className='row my-3'>
                <div className='col-md-12'>
                  <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                    <button onClick={onsubmit} className="submit-btn">
                      Submit
                    </button>
                    <a className="reset-btn" onClick={() => history.push("/")}>
                      Exit
                    </a>
                  </div>
                </div>
              </div>
          </div>
        </div>
      </div>
    </div>
    </div >
  );
};

export default EMDRequestFormEdit;

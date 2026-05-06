import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory } from 'react-router-dom';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import SPCRUDOPS from "../../service/BAL/spcrud";
import logo from "../../assets/sona-comstarlogo.png";

const Section = ({ title, children }: any) => (
  <div className="form-section">
    <h3>{title}</h3>
    {children}
  </div>
);

const Grid = ({ children }: any) => (
  <div className="form-grid">{children}</div>
);

const Field = ({ label, children, full }: any) => (
  <div className={full ? "form-field full" : "form-field"}>
    <label>{label}</label>
    {children}
  </div>
);

const EMDRequestForm = (props: ISonaEmdProps) => {
  const history = useHistory();
  const spCrudOps = SPCRUDOPS();

  // ---------- Dropdown options ----------
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
  const [vendorSiteKey, setVendorSiteKey] = useState<number | undefined>();
  const [productType, setProductType] = useState<number | undefined>();
  const [currency, setCurrency] = useState<number | undefined>();
  const [tenderType, setTenderType] = useState<number | undefined>(); // NEW
  // ---------- Other state ----------
  const [files, setFiles] = useState<File[]>([]);
  const [approvers, setApprovers] = useState<number[]>([]);
  const [approvalType, setApprovalType] = useState<any[]>([]);
  const [approverMatrix, setApproverMatrix] = useState<any[]>([]);

  const [employee, setEmployee] = React.useState({
    EmployeeCode: "",
    EmployeeName: "",
    Division: "",
    Department: "",
    Location: "",
    ReportingManager: "",
    HOD: "",
    ContactNo: "",
    EmployeeStatus: "",
    EmployeeEmail: "",
    ReportingManagerId: 0,
    HODId: 0

  });

  const [vendor, setVendor] = React.useState({
    VendorCode: "",
    VendorName: "",
    VendorSite: "",
    ContractType: "",
    TenderType: "",
    TenderNo: "",
    TenderDate: "",
    TenderAmount: "",
    EMDAmount: "",
    Currency: "",
    TenderClosingDate: "",
    EMDPercentage: "",
    ModeOfPayment: "",
    ProductType: "",
  });

  // In-memory cache to auto-fill vendor name/site by selected code
  const [vendorAll, setVendorAll] = useState<any[]>([]);
  const [wfHistory, setWfHistory] = useState<any[]>([]);
  const [commentHistory, setCommentHistory] = useState<any[]>([]);
  const [isSubmitting, setIsSubmitting] = useState(false);
  // ---------- Effects ----------
  useEffect(() => {
    getuserData();
    getCurrencyData();
    getVendorDataMaster();
    getContractTypeData();

    getModeOfPaymentData();
    getProductTypeData();
    getTenderTypeData(); // NEW

    loadApproverMatrix();
  }, []);

  // Auto-calc EMD %
  useEffect(() => {
    const tender = parseFloat(vendor.TenderAmount || "0");
    const emd = parseFloat(vendor.EMDAmount || "0");

    if (tender > 0) {
      const percent = ((emd / tender) * 100).toFixed(2);
      setVendor(prev => ({ ...prev, EMDPercentage: percent }));
    }
  }, [vendor.TenderAmount, vendor.EMDAmount]);


  useEffect(() => {

    if (!employee.ReportingManagerId || approvalType.length === 0) return;

    const rm = {
      Seq: 1,
      Role: "RM",
      ApproverID: employee.ReportingManagerId,
      ApproverName: employee.ReportingManager,
      Status: "Pending"
    };

    const hod = {
      Seq: 2,
      Role: "HOD",
      ApproverID: employee.HODId,
      ApproverName: employee.HOD,
      Status: "Waiting"
    };

    const others = approvalType.map((x: any, index: number) => ({
      Seq: index + 3,
      Role: x.Role?.RoleName,
      ApproverID: x.Approver?.ID,
      ApproverName: x.Approver?.Title,
      Status: "Waiting"
    }));

    setApproverMatrix([rm, hod, ...others]);

  }, [employee, approvalType]);
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
      const options = data.map((m: any) => ({
        key: m.Id,
        text: m.Currency
      }));
      setCurrencyOptions(options);
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

    // Cache full rows for auto-fill
    setVendorAll(res);

    setVendorNameOptions(res.map((v: any) => ({ key: v.ID, text: v.Name })));
    setVendorCodeOptions(res.map((v: any) => ({ key: v.ID, text: v.VendorCode })));
    setVendorSiteOptions(res.map((v: any) => ({ key: v.ID, text: v.VendorSite })));
  };

  const getuserData = async () => {
    (await spCrudOps).getData(
      "EmployeeMaster",
      "EmployeeCode,EmployeeName,Division,Location,Department,ReportingManager/EMail,ReportingManager/Id,HOD/EMail,HOD/Id,ReportingManager/Title,HOD/Title,Finance/Id,Finance/Title,Finance/EMail,ContactNo,EmployeeStatus,EmployeeEmail,Employee/Id,Employee/Title",
      "ReportingManager,HOD,Employee,Finance",
      `EmployeeId eq '${props.context.pageContext.legacyPageContext.userId}'`,
      { column: "ID", isAscending: true },
      1,
      props
    )
      .then((res: any[]) => {
        if (res.length > 0) {
          const userData = res[0];

          setEmployee({
            EmployeeCode: userData.EmployeeCode || "",
            EmployeeName: userData.EmployeeName || "",
            Division: userData.Division || "",
            Department: userData.Department || "",
            Location: userData.Location || "",
            ReportingManager: userData.ReportingManager?.Title || "",
            HOD: userData.HOD?.Title || "",
            ContactNo: userData.ContactNo || "",
            EmployeeStatus: userData.EmployeeStatus || "",
            EmployeeEmail: userData.EmployeeEmail || "",
            ReportingManagerId: userData.ReportingManager?.Id || 0,
            HODId: userData.HOD?.Id || 0,

          });

          const rmId = userData.ReportingManager?.Id;
          const hodId = userData.HOD?.Id;
          const userApprovers = [rmId, hodId].filter((id): id is number => !!id);

          setApprovers(prev => {
            const merged = [...prev, ...userApprovers];
            return merged.filter((value, index, self) => self.indexOf(value) === index);
          });
        } else {
          console.log("No user data found for the current user.");
        }
      })
      .catch((error: any) => {
        console.error("Error fetching user data:", error);
      });
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

    const options = res.map((m: any) => ({
      key: m.Id,
      text: m.Mode
    }));

    setModeOfPaymentOptions(options);
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

    const options = res.map((m: any) => ({
      key: m.Id,
      text: m.ProductType
    }));

    setProductTypeOptions(options);
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

    const options = res.map((c: any) => ({
      key: c.Id,
      text: c.ContractType
    }));

    setContractTypeOptions(options);
  };

  // NEW: Tender Type master
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

    const options = res.map((t: any) => ({
      key: t.Id,
      text: t.TenderType || t.Title
    }));

    setTenderTypeOptions(options);
  };

  // ---------- Helper: Duplicate tender number old ----------
  // const checkDuplicateTendorNumber = async (tendorNo: string): Promise<boolean> => {
  const checkDuplicateTendorNumber = async (tendorNo: string): Promise<any[]> => {
    const sp = await spCrudOps;
    const res = await sp.getData(
      "EMDDetails",
      "ID,TenderNo,Title",
      "",
      // `TenderNo eq '${tendorNo}'`,
      `TenderNo eq '${tendorNo}' and Status ne 'Rejected'`,

      { column: "ID", isAscending: true },
      100,
      props
    );
    // return res.length > 0;
    return res;
  };

  // ---------- Helpers: EMD Request No. ----------


  const getFinancialYear = (d = new Date()) => {
    const y = d.getFullYear();
    const m = d.getMonth() + 1; // 1..12
    if (m >= 4) {
      // Apr–Mar
      return `${y}-${String(y + 1).slice(-2)}`;
    } else {
      return `${y - 1}-${String(y).slice(-2)}`;
    }
    // Examples:
    // 10-Apr-2025 -> 2025-26
    // 10-Feb-2026 -> 2025-26
  };

  const pad = (n: number, width = 5) => String(n).padStart(width, "0");

  const generateEMDRequestNo = (id: number) => {
    const fy = getFinancialYear();
    return `EMD/${fy}/${pad(id)}`; // e.g., EMD2025-26_00023
  };

  //----------------------Get Approval Matrix---------------------------------
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

  // ---------- Submit ----------
  const onsubmit = async () => {
    if (isSubmitting) return;
    setIsSubmitting(true);
    // Date validation (must be greater than Tender Date)
    if (vendor.TenderDate && vendor.TenderClosingDate) {
      const tenderDate = new Date(vendor.TenderDate);
      const closingDate = new Date(vendor.TenderClosingDate);

      if (closingDate <= tenderDate) {
        alert("Tender Closing Date must be greater than Tender Date.");
        setIsSubmitting(false);
        return;
      }
    }
    try {
      const sp = await spCrudOps;

      // Validate
      if (!vendor.VendorCode && !vendorCodeKey) {
        alert("Vendor Code is required.");
        setIsSubmitting(false);
        return;
      }
      if (!contractType) {
        alert("Contract Type is required.");
        setIsSubmitting(false);
        return;
      }

      if (!vendor.TenderNo) {
        alert("Tender No. is required.");
        setIsSubmitting(false);
        return;
      }

      if (!vendor.TenderDate) {
        alert("Tender Date is required.");
        setIsSubmitting(false);
        return;
      }

      if (!tenderType) {
        alert("Tender Type is required.");
        setIsSubmitting(false);
        return;
      }

      if (!vendor.TenderAmount) {
        alert("Tender Amount is required.");
        setIsSubmitting(false);
        return;
      }
      if (!vendor.EMDAmount) {
        alert("EMD Amount is required.");
        setIsSubmitting(false);
        return;
      }
      if (!currency) {
        alert("Currency is required.");
        setIsSubmitting(false);
        return;
      }

      if (!vendor.TenderClosingDate) {
        alert("Tender Closing Date is required.");
        setIsSubmitting(false);
        return;
      }
      if (!modeOfPayment) {
        alert("Mode of Payment is required.");
        setIsSubmitting(false);
        return;
      }
      if (!productType) {
        alert("Product Type is required.");
        setIsSubmitting(false);
        return;
      }

      // ✅ Attachment validation (ADD THIS FIRST)
      if (!files || files.length === 0) {
        alert("Please upload at least one attachment.");
        setIsSubmitting(false);
        return;
      }

      // const isDuplicate = await checkDuplicateTendorNumber(vendor.TenderNo);
      // if (isDuplicate) {
      //   alert("❌ Tender Number already exists!");
      //   return;
      // }

      const isDuplicate = await checkDuplicateTendorNumber(vendor.TenderNo);
      // if (isDuplicate) {
      //   alert("⚠ Tender Number already exists in the system.\n!");
      // }
      if (isDuplicate.length > 0) {

        const message = isDuplicate
          .map((item: any, index: number) => {
            return `${index + 1}. ${item.Title}`;
          })
          .join("\n");

        alert(
          `⚠ Tender already exists!\n\nPrevious Requests:\n${message}`
        );
      }

      // ================= APPROVAL MATRIX BUILD =================

      // 🔹 RM (from EmployeeMaster)
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

      // 🔥 INITIAL WF HISTORY
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
          ActionTaken: "Submitted",
          Comment: "Request Created",
          Date: formattedDate,
          CurrentStatus: "Pending for Approval"
        }
      ];

      // Insert Parent
      const parentResponse = await sp.insertData(
        "EMDDetails",
        {
          // ----- Requestor -----
          EmployeeName: employee.EmployeeName,
          EmployeeCode: employee.EmployeeCode,
          Division: employee.Division,
          Location: employee.Location,
          RM: employee.ReportingManager,
          HOD: employee.HOD,
          ContactNo: employee.ContactNo,
          EmployeeStatus: employee.EmployeeStatus,
          Department: employee.Department,
          Email: employee.EmployeeEmail,

          // ----- Vendor -----
          VendorCode: vendor.VendorCode || "",             // text
          VendorNameId: Number(vendorName),                // lookup id to VendorMaster
          VendorSite: vendor.VendorSite || "",             // text

          // ----- Masters -----
          ContractTypeId: Number(contractType),
          ModeofPaymentId: Number(modeOfPayment),
          ProductTypeId: Number(productType),
          CurrencyId: Number(currency),
          TenderTypeId: Number(tenderType),                // NEW

          // ----- Tender -----
          TenderNo: vendor.TenderNo,
          TenderDate: vendor.TenderDate,
          TenderAmount: vendor.TenderAmount,
          EMDAmount: vendor.EMDAmount,
          TenderClosingDate: vendor.TenderClosingDate,
          EMDPercentage: vendor.EMDPercentage,

          // ----- Approval Matrix -----
          ApprovalMatrix: JSON.stringify(approvalMatrix),
          CurrentApproverId: approvalMatrix[0]?.ApproverID,
          PendingAt: approvalMatrix[0]?.Role,
          RequestorStatus: "Pending For Approval",
          ApproverStatus: "Pending",
          Status: "Pending for Approval",
          DuplicateTender: isDuplicate ? true : false,
          WFHistory: JSON.stringify(initialWFHistory),
          CommentHistory: JSON.stringify(initialWFHistory),
        },
        props
      );

      const requestId = parentResponse.data.ID;
      console.log("✅ Parent Saved ID:", requestId);

      // Generate EMD Request No and update Title immediately
      const emdNo = generateEMDRequestNo(requestId);
      await sp.updateData(
        "EMDDetails",
        requestId,
        { Title: emdNo }, // Title stores EMDRequestNo
        props
      );

      // Attachments (if any) old for single attachemnt
      // if (files.length > 0) {
      //   await Promise.all(
      //     files.map(file =>
      //       sp.addAttchmentInList(file, "EMDDetails", requestId, file.name, props)
      //     )
      //   );
      // if (files.length > 0) {
      //   for (const file of files) {
      //     await sp.addAttchmentInList(
      //       file,
      //       "EMDDetails",
      //       requestId,
      //       file.name,
      //       props
      //     );
      //   }
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
            requestId,
            uniqueName,
            props
          );
        }
      }

      alert(`✅ Submitted! Your EMD Request No. is ${emdNo}`);
      history.push("/");

    } catch (error) {
      console.error("❌ Error submitting data:", error);
      alert("Something went wrong. Please check console.");
      setIsSubmitting(false);
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
                <h1>New EMD Issuance </h1>
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
                      <label className='font'>Vendor Code <span className='Mantorystar'>*</span></label>
                      <Dropdown
                        className='formtext-control'
                        placeholder="Select Vendor Code"
                        options={vendorCodeOptions}
                        selectedKey={vendorCodeKey}
                        onChange={(e, option) => {
                          if (option) {
                            setVendorCodeKey(option.key as number);
                            setVendor(prev => ({ ...prev, VendorCode: option.text }));

                            // Auto-fill Name & Site from selected row
                            const row = vendorAll.find((v: any) => v.ID === option.key);
                            if (row) {
                              // Name (lookup id)
                              setVendorName(row.ID);

                              // Site (text + selected key match)
                              setVendorSiteKey(row.ID);
                              setVendor(prev => ({ ...prev, VendorSite: row.VendorSite }));
                            }
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
                        disabled
                        placeholder="Auto-filled from Vendor Code"
                      />
                    </div>
                    <div className='col-md-4'>
                      <label className='font'>Vendor Site </label>
                      <Dropdown
                        className='formtext-control'
                        options={vendorSiteOptions}
                        selectedKey={vendorSiteKey}
                        disabled
                        placeholder="Auto-filled from Vendor Code"
                      />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">Contract Type <span className='Mantorystar'>*</span></label>
                      <Dropdown
                        className='formtext-control'
                        options={contractTypeOptions}
                        selectedKey={contractType}
                        onChange={(e, option) => {
                          if (option) setContractType(option.key as number);
                        }}
                        placeholder="Select Contract Type"
                      />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Tender No <span className='Mantorystar'>*</span></label>
                      <input
                        className='form-control'
                        value={vendor.TenderNo}
                        onChange={(e) => {
                          const value = e.target.value;
                          // Only alphanumeric (no special characters)
                          if (/^[a-zA-Z0-9]*$/.test(value)) {
                            setVendor(prev => ({ ...prev, TenderNo: value }));
                          } else {
                            alert("Special characters are not allowed");
                          }
                        }}
                      />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Tender Date <span className='Mantorystar'>*</span></label>
                      <input
                        className='form-control'
                        type="date"
                        value={vendor.TenderDate}
                        onChange={(e) => setVendor(prev => ({ ...prev, TenderDate: e.target.value }))}
                      />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">Tender Type <span className='Mantorystar'>*</span></label>
                      <Dropdown
                        className='formtext-control'
                        options={tenderTypeOptions}
                        selectedKey={tenderType}
                        onChange={(e, option) => {
                          if (option) setTenderType(option.key as number);
                        }}
                        placeholder="Select Tender Type"
                      />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Tender Amount <span className='Mantorystar'>*</span></label>
                      <input
                        className="form-control"
                        value={vendor.TenderAmount}
                        onChange={(e) => {
                          const value = e.target.value;

                          // Sirf number allow karega
                          if (/^\d*$/.test(value)) {
                            setVendor(prev => ({ ...prev, TenderAmount: value }));
                          } else {
                            alert("Please type number only");
                          }
                        }}
                      />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">EMD Amount <span className='Mantorystar'>*</span></label>
                      <input
                        className="form-control"
                        value={vendor.EMDAmount}
                        onChange={(e) => {
                          const value = e.target.value;

                          if (/^\d*$/.test(value)) {
                            setVendor(prev => ({ ...prev, EMDAmount: value }));
                          } else {
                            alert("Please type number only");
                          }
                        }}
                      />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">Currency <span className='Mantorystar'>*</span></label>
                      <Dropdown
                        className='formtext-control'
                        options={currencyOptions}
                        selectedKey={currency}
                        onChange={(e, option) => {
                          if (option) setCurrency(option.key as number);
                        }}
                        placeholder="Select Currency"
                      />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Tender Closing Date <span className='Mantorystar'>*</span></label>
                      <input
                        type="date" className="form-control"
                        value={vendor.TenderClosingDate}
                        onChange={(e) => setVendor(prev => ({ ...prev, TenderClosingDate: e.target.value }))}
                      />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">EMD Percentage <span className='Mantorystar'>*</span></label>
                      <input
                        value={vendor.EMDPercentage} className="form-control"
                        onChange={(e) => setVendor(prev => ({ ...prev, EMDPercentage: e.target.value }))}
                        readOnly
                      />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">Mode of Payment <span className='Mantorystar'>*</span></label>
                      <Dropdown
                        className='formtext-control'
                        options={modeOfPaymentOptions}
                        selectedKey={modeOfPayment}
                        onChange={(e, option) => {
                          if (option) setModeOfPayment(option.key as number);
                        }}
                        placeholder="Select Mode of Payment"
                      />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Product Type <span className='Mantorystar'>*</span></label>
                      <Dropdown
                        className='formtext-control'
                        options={productTypeOptions}
                        selectedKey={productType}
                        onChange={(e, option) => {
                          if (option) setProductType(option.key as number);
                        }}
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
                      <label className='font'>
                        Attachment(s) <span className='Mantorystar'>*</span>
                      </label>
                      <input type="file" multiple
                        onChange={(e) => {
                          if (e.target.files) {
                            setFiles(Array.from(e.target.files));
                          }
                        }}
                      />
                    </div>
                  </div>
                </div>
                <div className='row my-3'>
                  <div className='col-md-12'>
                    <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                      <button onClick={onsubmit} disabled={isSubmitting} className="submit-btn">
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
      </div>
  );
};

export default EMDRequestForm;
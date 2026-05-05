import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory } from 'react-router-dom';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import SPCRUDOPS from "../../service/BAL/spcrud";

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
    <div className="forex-wrapper">

      {/* ================= HEADER ================= */}
      <div className="forex-header">
        <h2>New EMD Issuance</h2>
      </div>

      <div className="forex-card">

        {/* ================= APPROVAL HIERARCHY ================= */}
        {/* <div className="approval-hierarchy">
          <div className="step">
            <div className="step-title">{employee.EmployeeName || "Employee"}</div>
          </div>

          <div className="arrow">→</div>

          <div className="step">
            <div className="step-title">{employee.ReportingManager || "Reporting Manager"}</div>
          </div>

          <div className="arrow">→</div>

          <div className="step">
            <div className="step-title">{employee.HOD || "HOD"}</div>
          </div>
        </div> */}

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
        </div>
        {/* ================= REQUESTOR ================= */}
        <Section title="Requestor Information">
          <Grid style={{ marginTop: "20px" }}>
            <Field label="Employee Code">
              <input type="text" value={employee.EmployeeCode} readOnly />
            </Field>

            <Field label="Employee Name">
              <input type="text" value={employee.EmployeeName} readOnly />
            </Field>

            <Field label="Division">
              <input type="text" value={employee.Division} readOnly />
            </Field>

            <Field label="Location">
              <input type="text" value={employee.Location} readOnly />
            </Field>

            <Field label="Reporting Manager">
              <input type="text" value={employee.ReportingManager} readOnly />
            </Field>

            <Field label="HOD">
              <input type="text" value={employee.HOD} readOnly />
            </Field>

            <Field label="Contact No">
              <input type="text" value={employee.ContactNo} readOnly />
            </Field>

            <Field label="Employee Status">
              <input type="text" value={employee.EmployeeStatus} readOnly />
            </Field>

            <Field label="Department">
              <input type="text" value={employee.Department} readOnly />
            </Field>

            <Field label="Employee Email" full>
              <input type="email" value={employee.EmployeeEmail} readOnly />
            </Field>
          </Grid>
        </Section>

        {/* ================= VENDOR / REQUEST ================= */}
        <Section title="EMD Request Details">
          <Grid>
            {/* Vendor Code (drives Name & Site) */}
            <Field
              //  label="Vendor Code"
              label={
                <span>
                  Vendor Code <span style={{ color: "red" }}>*</span>
                </span>
              }
            >
              <Dropdown
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
                placeholder="Select Vendor Code"
              />
            </Field>

            {/* Vendor Name (auto-filled) */}
            <Field label="Vendor Name">
              <Dropdown
                options={vendorNameOptions}
                selectedKey={vendorName}
                disabled
                placeholder="Auto-filled from Vendor Code"
              />
            </Field>

            {/* Vendor Site (auto-filled) */}
            <Field label="Vendor Site">
              <Dropdown
                options={vendorSiteOptions}
                selectedKey={vendorSiteKey}
                disabled
                placeholder="Auto-filled from Vendor Code"
              />
            </Field>



            <Field
              // label="Contract Type"
              label={
                <span>
                  Contract Type <span style={{ color: "red" }}>*</span>
                </span>
              }
            >
              <Dropdown
                options={contractTypeOptions}
                selectedKey={contractType}
                onChange={(e, option) => {
                  if (option) setContractType(option.key as number);
                }}
                placeholder="Select Contract Type"
              />
            </Field>

            {/* <Field label="Tender No.">
              <input
                value={vendor.TenderNo}
                onChange={(e) => setVendor(prev => ({ ...prev, TenderNo: e.target.value }))}
              />
            </Field> */}
            <Field
              // label="Tender No."
              label={
                <span>
                  Tender No. <span style={{ color: "red" }}>*</span>
                </span>
              }
            >
              <input
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
            </Field>

            <Field
              //  label="Tender Date"
              label={
                <span>
                  Tender Date <span style={{ color: "red" }}>*</span>
                </span>
              }
            >
              <input
                type="date"
                value={vendor.TenderDate}
                onChange={(e) => setVendor(prev => ({ ...prev, TenderDate: e.target.value }))}
              />
            </Field>

            {/* ✅ Tender Type as Dropdown from master */}
            <Field
              // label="Tender Type"
              label={
                <span>
                  Tender Type <span style={{ color: "red" }}>*</span>
                </span>
              }
            >
              <Dropdown
                options={tenderTypeOptions}
                selectedKey={tenderType}
                onChange={(e, option) => {
                  if (option) setTenderType(option.key as number);
                }}
                placeholder="Select Tender Type"
              />
            </Field>

            {/* <Field label="Tender Amount">
              <input
                value={vendor.TenderAmount}
                onChange={(e) => setVendor(prev => ({ ...prev, TenderAmount: e.target.value }))}
              />
            </Field> */}
            <Field
              // label="Tender Amount"
              label={
                <span>
                  Tender Amount <span style={{ color: "red" }}>*</span>
                </span>
              }

            >
              <input
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
            </Field>

            {/* <Field label="EMD Amount">
              <input
                value={vendor.EMDAmount}
                onChange={(e) => setVendor(prev => ({ ...prev, EMDAmount: e.target.value }))}
              />
            </Field> */}

            <Field
              //  label="EMD Amount"
              label={
                <span>
                  EMD Amount <span style={{ color: "red" }}>*</span>
                </span>
              }

            >
              <input
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
            </Field>

            <Field
              // label="Currency"
              label={
                <span>
                  Currency <span style={{ color: "red" }}>*</span>
                </span>
              }
            >
              <Dropdown
                options={currencyOptions}
                selectedKey={currency}
                onChange={(e, option) => {
                  if (option) setCurrency(option.key as number);
                }}
                placeholder="Select Currency"
              />
            </Field>

            <Field
              //  label="Tender Closing Date"
              label={
                <span>
                  Tender Closing Date <span style={{ color: "red" }}>*</span>
                </span>
              }

            >
              <input
                type="date"
                value={vendor.TenderClosingDate}
                onChange={(e) => setVendor(prev => ({ ...prev, TenderClosingDate: e.target.value }))}
              />
            </Field>

            <Field
              //  label="EMD Percentage"
              label={
                <span>
                  EMD Percentage <span style={{ color: "red" }}>*</span>
                </span>
              }

            >
              <input
                value={vendor.EMDPercentage}
                onChange={(e) => setVendor(prev => ({ ...prev, EMDPercentage: e.target.value }))}
                readOnly
              />
            </Field>

            <Field
              // label="Mode Of Payment"  
              label={
                <span>
                  Mode of Payment <span style={{ color: "red" }}>*</span>
                </span>
              }
            >
              <Dropdown
                options={modeOfPaymentOptions}
                selectedKey={modeOfPayment}
                onChange={(e, option) => {
                  if (option) setModeOfPayment(option.key as number);
                }}
                placeholder="Select Mode of Payment"
              />
            </Field>

            <Field
              //  label="Product Type"
              label={
                <span>
                  Product Type <span style={{ color: "red" }}>*</span>
                </span>
              }

            >
              <Dropdown
                options={productTypeOptions}
                selectedKey={productType}
                onChange={(e, option) => {
                  if (option) setProductType(option.key as number);
                }}
                placeholder="Select Product Type"
              />
            </Field>
          </Grid>
        </Section>

        <Section
          // title="Upload Documents"
          title={
            <span>
              Upload Documents <span style={{ color: "red" }}>*</span>
            </span>
          }

        >
          <Grid>
            <Field label="Attachment(s)">
              <input
                type="file"
                multiple
                onChange={(e) => {
                  if (e.target.files) {
                    setFiles(Array.from(e.target.files));
                  }
                }}
              />
            </Field>
          </Grid>
        </Section>

        <div className="button-row">
          <button className="btn-submit" onClick={onsubmit} disabled={isSubmitting}>Submit</button>
          <button className="btn-exit" onClick={() => history.push("/")}>Exit</button>
        </div>

      </div>

    </div>
  );
};

export default EMDRequestForm;
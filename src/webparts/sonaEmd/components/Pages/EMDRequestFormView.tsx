import React, { useEffect, useMemo, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useHistory, useLocation } from 'react-router-dom';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import SPCRUDOPS from "../../service/BAL/spcrud";

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

const EMDRequestFormView = (props: ISonaEmdProps) => {
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

  // ---------- Misc ----------
  const [files, setFiles] = useState<File[]>([]);
  const [title, setTitle] = useState<string>(""); // EMD Request No. (Title)

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
  }, []);

  // ---------- After masters ready, load the item by id ----------
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
        },
        props
      );

      // Attachments: add any new ones
      if (files.length > 0) {
        await Promise.all(
          files.map(file => sp.addAttchmentInList(file, "EMDDetails", itemId, file.name, props))
        );
      }

      alert("✅ Request updated successfully!");
      history.push("/");

    } catch (error) {
      console.error("❌ Error updating data:", error);
      alert("Something went wrong. Please check console.");
    }
  };

  return (
    <div className="forex-wrapper">

      {/* ================= HEADER ================= */}
      <div className="forex-header">
        <h2>Edit EMD Issuance</h2>
      </div>

      <div className="forex-card">

        {/* (Optional) EMD Request No. (Title) */}
        {/* {title ? (
          <Section title="EMD Request No.">
            <Grid>
              <Field label="EMD Request No.">
                <input type="text" value={title} readOnly />
              </Field>
            </Grid>
          </Section>
        ) : null} */}

                {/* ================= APPROVAL HIERARCHY ================= */}
        <div className="emd-hierarchy">
          <div className="emd-step active-step">{employee.EmployeeName}</div>

          <div className="emd-step"  style={{marginLeft:"30px"}}>{employee.ReportingManager}</div>

          <div className="emd-step" style={{marginLeft:"30px"}}>{employee.HOD}</div>
        </div>

        {/* ================= REQUESTOR ================= */}
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

        {/* ================= VENDOR ================= */}
        <Section title="EMD Request Details">
          <Grid>

            {/* Vendor Code (changes should drive name & site auto-fill) */}
            <Field label="Vendor Code">
              <Dropdown
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
                disabled
                placeholder="Select Vendor Code"
              />
            </Field>

            {/* Vendor Name (auto-filled from selection, लेकिन editable छोड़ रहा हूँ; चाहें तो disabled कर दें) */}
            <Field label="Vendor Name">
              <Dropdown
                options={vendorNameOptions}
                selectedKey={vendorName}
                onChange={(e, option) => {
                  if (option) setVendorName(option.key as number);
                }}
                disabled
                placeholder="Select Vendor Name"
              />
            </Field>

            {/* Vendor Site (text match) */}
            <Field label="Vendor Site">
              <Dropdown
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
            </Field>



            <Field label="Contract Type">
              <Dropdown
                options={contractTypeOptions}
                selectedKey={contractType}
                onChange={(e, option) => option && setContractType(option.key as number)}
                placeholder="Select Contract Type"
                disabled
              />
            </Field>

            <Field label="Tender No.">
              <input
                value={vendor.TenderNo}
                onChange={(e) => setVendor(prev => ({ ...prev, TenderNo: e.target.value }))}
                disabled
              />
            </Field>

            <Field label="Tender Date">
              <input
                type="date"
                value={vendor.TenderDate}
                onChange={(e) => setVendor(prev => ({ ...prev, TenderDate: e.target.value }))}
                disabled
              />
            </Field>

            {/* Tender Type */}
            <Field label="Tender Type">
              <Dropdown
                options={tenderTypeOptions}
                selectedKey={tenderType}
                onChange={(e, option) => option && setTenderType(option.key as number)}
                disabled
                placeholder="Select Tender Type"
              />
            </Field>

            <Field label="Tender Amount">
              <input
                value={vendor.TenderAmount}
                onChange={(e) => setVendor(prev => ({ ...prev, TenderAmount: e.target.value }))}
                disabled
              />
            </Field>

            <Field label="EMD Amount">
              <input
                value={vendor.EMDAmount}
                onChange={(e) => setVendor(prev => ({ ...prev, EMDAmount: e.target.value }))}
                disabled
              />
            </Field>

            <Field label="Currency">
              <Dropdown
                options={currencyOptions}
                selectedKey={currency}
                onChange={(e, option) => option && setCurrency(option.key as number)}
                placeholder="Select Currency"
                disabled
              />
            </Field>

            <Field label="Tender Closing Date">
              <input
                type="date"
                value={vendor.TenderClosingDate}
                onChange={(e) => setVendor(prev => ({ ...prev, TenderClosingDate: e.target.value }))}
                disabled
              />
            </Field>

            <Field label="EMD Percentage">
              <input
                value={vendor.EMDPercentage}
                onChange={(e) => setVendor(prev => ({ ...prev, EMDPercentage: e.target.value }))}
                readOnly
                disabled
              />
            </Field>

            <Field label="Mode Of Payment">
              <Dropdown
                options={modeOfPaymentOptions}
                selectedKey={modeOfPayment}
                onChange={(e, option) => option && setModeOfPayment(option.key as number)}
                placeholder="Select Mode of Payment"
                disabled
              />
            </Field>

            <Field label="Product Type">
              <Dropdown
                options={productTypeOptions}
                selectedKey={productType}
                onChange={(e, option) => option && setProductType(option.key as number)}
                placeholder="Select Product Type"
                disabled
              />
            </Field>
          </Grid>
        </Section>

        <Section title="Upload Documents">
          <Grid>
            <Field label="Attachment">
              <input
                type="file"
                multiple
                onChange={(e) => {
                  if (e.target.files) setFiles(Array.from(e.target.files));
                }}
              />
            </Field>
          </Grid>
        </Section>

        <div className="button-row">
          {/* <button className="btn-submit" onClick={onsubmit}>Submit</button> */}
          <button className="btn-exit" onClick={() => history.push("/")}>Exit</button>
        </div>

      </div>

    </div>
  );
};

export default EMDRequestFormView;

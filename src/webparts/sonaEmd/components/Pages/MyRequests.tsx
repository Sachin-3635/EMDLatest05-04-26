import * as React from "react";
import "./Css/InitiatorDashboard.scss";
import { ISonaEmdProps } from "../ISonaEmdProps";
import { useState, useEffect, useMemo } from "react";
import { Link } from "react-router-dom";
import Left from "../../assets/LeftArrow.png";
import Right from "../../assets/RightArrow.png";
import View from "../../assets/Eye.png";
import Edit from "../../assets/Pencil.png"
import SPCRUDOPS from "../../service/BAL/spcrud";


// type Lookup<T = string> = { Title?: T; Currency?: T; TenderType?: T };
type Lookup<T = string> = { Name?: T; Currency?: T; TenderType?: T };

interface EmdItem {
  Id: number;                 // ✅ REAL SharePoint item id
  Title: string;              // ✅ EMD Request No. (Title is text)
  EmployeeName?: string;
  TenderNo?: string;
  TenderDate?: string;
  TenderAmount?: string;
  EMDAmount?: string;
  EMDPercentage?: string;
  TenderClosingDate?: string;

  // Lookups
  VendorName?: Lookup<string>;
  TenderType?: Lookup<string>;
  Currency?: Lookup<string>;

  // System
  Modified?: string;
  Created?: string;
  AuthorId?: number;
  Status?: string;
}

export const MyRequests: React.FC<ISonaEmdProps> = (props: ISonaEmdProps) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [listData, setListData] = useState<EmdItem[]>([]);
  const [filteredData, setFilteredData] = useState<EmdItem[]>([]);
  const [loading, setLoading] = useState(true);

  // Pagination
  const itemsPerPage = 10;
  const [currentPage, setCurrentPage] = useState(1);

  const formatDate = (dateString?: string) => {
    if (!dateString) return "-";
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return "-";
    return date
      .toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" })
      .replace(/\//g, "-");
  };

  const formatAmount = (value?: string | number) => {
    if (value === undefined || value === null || value === "") return "-";
    const num = typeof value === "number" ? value : Number(value);
    if (isNaN(num)) return String(value);
    return num.toLocaleString("en-IN");
  };

  // 🔹 Fetch Data
  const GetListData = async () => {
    try {
      setLoading(true);
      const spCrudOps = await SPCRUDOPS();

      const selectFields = [
        "Id",
        "Title",
        "EmployeeName",
        "TenderNo",
        "TenderDate",
        "TenderAmount",
        "EMDAmount",
        "EMDPercentage",
        "TenderClosingDate",
        "VendorNameId",
        "VendorName/Title",
        "VendorName/Name",
        "TenderTypeId",
        "TenderType/TenderType",
        "CurrencyId",
        "Currency/Currency",
        "Modified",
        "Created",
        "AuthorId",
        "Status",
      ].join(",");

      const expandFields = "VendorName,TenderType,Currency";

      // ✅ Filter by current user (Created By)
      const currentSpUserId = props?.context?.pageContext?.legacyPageContext?.userId;
      const filter = Number.isFinite(currentSpUserId) ? `AuthorId eq ${currentSpUserId}` : "";

      const parentItems: EmdItem[] = await spCrudOps.getData(
        "EMDDetails",
        selectFields,
        expandFields,
        filter,
        { column: "Id", isAscending: false },   // ✅ latest first
        5000,
        props
      );

      const safeItems = Array.isArray(parentItems) ? parentItems : [];
      setListData(safeItems);
      setFilteredData(safeItems);
    } catch (error) {
      console.error("GetListData error:", error);
      setListData([]);
      setFilteredData([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    GetListData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  //  Search (only on existing fields)
  // useEffect(() => {
  //   let data = Array.isArray(listData) ? [...listData] : [];

  //   const q = searchTerm.trim().toLowerCase();
  //   if (q) {
  //     data = data.filter((item) => {
  //       const haystack = [
  //         item?.TenderNo,
  //         item?.EmployeeName,
  //         item?.VendorName?.Title,
  //         item?.TenderType?.TenderType,
  //         item?.EMDAmount,
  //         item?.Currency?.Currency,
  //         item?.Title, // EMD Request No.
  //       ]
  //         .filter(Boolean)
  //         .join(" ")
  //         .toLowerCase();

  //       return haystack.includes(q);
  //     });
  //   }

  //   setFilteredData(data);
  //   setCurrentPage(1);
  // }, [searchTerm, listData]);

  useEffect(() => {
    let data = [...listData];

    const q = searchTerm.trim().toLowerCase();

    if (q) {
      data = data.filter((item) => {
        const haystack = [
          item?.TenderNo,
          item?.EmployeeName,
          // item?.VendorName?.Title,
          item?.VendorName?.Name, //  also search by Vendor's internal Name (not just Title)
          item?.TenderType?.TenderType,
          item?.EMDAmount,
          item?.Currency?.Currency,
          item?.Title
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();

        return haystack.includes(q);
      });
    }

    //  Status Filter Apply
    if (statusFilter) {
      data = data.filter(
        (item) => (item?.Status || "Pending for Approval") === statusFilter
      );
    }

    setFilteredData(data);
    setCurrentPage(1);
  }, [searchTerm, statusFilter, listData]);


  // 🔹 Pagination
  const totalPages = Math.ceil((filteredData?.length || 0) / itemsPerPage);

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) setCurrentPage(page);
  };

  // ✅ No need to sort by Title (string). Sort by Id desc for stability.
  const sortedData = useMemo(
    () => [...(filteredData || [])].sort((a, b) => (b?.Id || 0) - (a?.Id || 0)),
    [filteredData]
  );

  const paginatedData = sortedData.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );


  const getStatusClass = (status?: string) => {
    switch ((status || "Pending for Approval").toLowerCase()) {
      case "pending for vouching": return "badge badge-vouching";
      case "pending for payment": return "badge badge-payment";
      default: return "badge badge-approval";
    }
  };
  return (
    <div className="dashboard-wrapper">
      {/* Header */}
      <div className="header">
        <div className="left-banner">
          <div className="logo-text">
            <h2>EMD Initiator</h2>
          </div>
        </div>
      </div>
      <div className='col-md-12 px-2 py-2 d-flex justify-content-between align-items-center flex-wrap'>
        <div className=" d-flex justify-content-between align-items-center" style={{ gap: "5px" }}>
          <input type="text" placeholder="Search by Tender No, Vendor, Amount..."
            className="form-control" style={{ width: "250px" }}
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
          <select
            className="form-control"
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            style={{ width: "250px" }}
          >
            <option value="">All Status</option>
            <option value="Pending for Approval">Pending for Approval</option>
            <option value="Pending for Vouching">Pending for Vouching</option>
            <option value="Pending for Payment">Pending for Payment</option>
            <option value="Closed">Closed</option>
            <option value="Rejected">Rejected</option>
            <option value="EMD Paid">EMD Paid</option>
          </select>
        </div>
        <div className='Dashbaordcreatebutton'>
          <Link to="/EMDRequestForm" className="create-button"> Request for EMD </Link>
          {/* <Link to="/NewRequest" className='create-button'>Deviation Request</Link> */}
        </div>
      </div>
      <main className="Main-Dash mx-2">
        <div className="overflow-x-auto">
          <div className="table-vert-scroll">

            <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
              <thead
                style={{ backgroundColor: "#3c3e45" }}
                className="text-white"
              >
                <tr>
                  <th className="px-4 py-2">EMD Request No.</th>
                  <th className="px-4 py-2">Tender No.</th>
                  <th className="px-4 py-2">Tender Type</th>
                  <th className="px-4 py-2">Customer/Vendor</th>
                  <th className="px-4 py-2">Employee Name</th>
                  <th className="px-4 py-2">EMD Amount</th>
                  <th className="px-4 py-2">Currency</th>
                  <th className="px-4 py-2">Status</th>
                  <th className="px-4 py-2">View</th>
                </tr>
              </thead>
              <tbody>
                {loading ? (
                  <tr>
                    <td colSpan={9} className="text-center">
                      Loading data...
                    </td>
                  </tr>
                ) : paginatedData.length === 0 ? (
                  <tr>
                    <td colSpan={9} className="text-center">
                      No records found
                    </td>
                  </tr>
                ) : (
                  paginatedData.map((item) => (
                    <tr key={item.Id}>
                      <td className="px-4 py-2">{item?.Title ?? "-"}</td>
                      <td className="px-4 py-2">{item?.TenderNo ?? "-"}</td>
                      <td className="px-4 py-2">{item?.TenderType?.TenderType ?? "-"}</td>
                      {/* <td className="px-4 py-2">{item?.VendorName?.Title ?? "-"}</td> */}
                      <td className="px-4 py-2">{item?.VendorName?.Name ?? "-"}</td>
                      <td className="px-4 py-2">{item?.EmployeeName ?? "-"}</td>
                      <td className="px-4 py-2">{item?.EMDAmount ? `₹ ${formatAmount(item.EMDAmount)}` : "-"}</td>
                      <td className="px-4 py-2">{item?.Currency?.Currency ?? "-"}</td>
                      {/* <td className="px-4 py-2">{formatDate(item?.TenderClosingDate)}</td> */}
                      <td className="px-4 py-2">
                        {item?.Status || "Pending for Approval"}
                        {/* <span className={getStatusClass(item?.Status)}>

                        </span> */}
                      </td>
                      <td className="px-4 py-2">
                        {/*  Use real Id here so EMDRequestFormView can prefill */}
                        {/* <Link to={`/EMDRequestFormView?ItemId=${item.Id}`}> */}
                        <Link to={`/ViewForm?ItemId=${item.Id}`}>
                          <img src={View} width={16} alt="View" />
                        </Link>
                        <Link to={`/EMDRequestFormEdit?ItemId=${item.Id}`}>
                          {/* <img src={Edit} width={16} alt="Edit" /> */}
                        </Link>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>

          {/* Pagination */}
          <div className="flex justify-center mt-6 overflow-x-auto">
            <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-#2149d5 rounded shadow" style={{ textAlign: "end" }}>
              {/* Previous Button */}
              <button
                onClick={() => handlePageChange(currentPage - 1)}
                disabled={currentPage === 1}
                style={{
                  backgroundColor: "#fff",
                  border: "1px solid #000 !important",
                  marginRight: "5px",
                  opacity: currentPage === 1 ? 0.5 : 1,
                }}
                className="px-3 py-1 border rounded"
              >
                <img src={Left} alt="" width={15} />
              </button>
              {/* Main Page Numbers */}
              {Array.from({ length: totalPages }, (_, i) => i + 1)
                .filter((page) => Math.abs(page - currentPage) <= 2)
                .map((page) => (
                  <button
                    key={page}
                    onClick={() => handlePageChange(page)}
                    style={{
                      backgroundColor: currentPage === page ? "#3c3e45" : "#fff",
                      color: currentPage === page ? "#fff" : "#000",
                      fontWeight: currentPage === page ? "bold" : "normal",
                      margin: currentPage === page ? "5px" : "5px",
                    }}
                    className="px-3 py-1 border rounded"
                  >
                    {page}
                  </button>
                ))}



              {/* Next Button */}
              <button
                onClick={() => handlePageChange(currentPage + 1)}
                disabled={currentPage === totalPages}
                style={{
                  backgroundColor: "#fff",
                  border: "1px solid #000 !important",
                  marginLeft: "5px",
                  opacity: currentPage === totalPages ? 0.5 : 1,
                }}
                className="px-3 py-1 border rounded"
              >
                <img src={Right} alt="" width={15} />
              </button>
            </div>
          </div>
        </div>
      </main>



    </div>
  );
};

export default MyRequests;
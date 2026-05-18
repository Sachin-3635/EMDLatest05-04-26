import * as React from "react";
import { useState, useEffect, useMemo } from "react";
import { useHistory, Link } from "react-router-dom";
import type { ISonaEmdProps } from "../ISonaEmdProps";
import "../../components/Pages/Css/InitiatorDashboard.scss";
import SPCRUDOPS from "../../service/BAL/spcrud";

import Left from "../../assets/LeftArrow.png";
import Right from "../../assets/RightArrow.png";
import View from "../../assets/Eye.png";
import Edit from "../../assets/Pencil.png";
import Renew from "../../assets/Renew.png";

type Lookup<T = string> = { Name?: T; Title?: T; Currency?: T; TenderType?: T };

interface EmdItem {
  Id: number;
  Title?: string;           // EMD Request No. (Title)
  TenderNo?: string;
  EMDAmount?: string | number;
  Status?: string;
  Author?: Lookup<string>;

  // Lookups
  VendorName?: Lookup<string>;
  TenderType?: Lookup<string>;
  Currency?: Lookup<string>;

  // For optional "Assigned to me" filter (stored as text in your earlier code)
  RM?: string;
  HOD?: string;

  // System
  Created?: string;
  Modified?: string;
}

export const ApprovalDashboard: React.FC<ISonaEmdProps> = (props: ISonaEmdProps) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState<"All" | "Pending for Approval" | "Pending for Vouching" | "Pending for Payment">("All");
  const [listData, setListData] = useState<EmdItem[]>([]);
  const [filteredData, setFilteredData] = useState<EmdItem[]>([]);
  const [loading, setLoading] = useState(true);

  // Pagination
  const itemsPerPage = 10;
  const [currentPage, setCurrentPage] = useState(1);

  const history = useHistory();

  const formatDate = (dateString?: string) => {
    if (!dateString) return "-";
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return "-";
    return date
      .toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
      })
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

      // Select the needed fields
      const selectFields = [
        "Id",
        "Title",
        "Author/Title",
        "TenderNo",
        "EMDAmount",
        "Status",                  // Choice column
        "RM", "HOD",              // For "Assigned to me" filter (if you want)
        "TenderTypeId",
        "TenderType/TenderType",
        "VendorNameId",
        "VendorName/Title",
        "VendorName/Name",
        "CurrencyId",
        "Currency/Currency",
        "Created",
        "Modified",
      ].join(",");

      const expandFields = "VendorName,TenderType,Currency,Author";

      // Server-side filter (optional): You can fetch all and filter client-side, which is simpler.
      // If you want to fetch only those with any status of interest, you could do:
      // const filter = "(Status eq 'Pending for Approval') or (Status eq 'Pending for Vouching') or (Status eq 'Pending for Payment')";
      const currentUserId = props.context.pageContext.legacyPageContext.userId;

      const filter = `Status eq 'Pending for Approval' and CurrentApproverId eq ${currentUserId}`; // fetch all, then filter client-side

      const items: EmdItem[] = await spCrudOps.getData(
        "EMDDetails",
        selectFields,
        expandFields,
        filter,
        { column: "Id", isAscending: false },  // latest first
        5000,
        props
      );

      const safeItems = Array.isArray(items) ? items : [];
      setListData(safeItems);
      setFilteredData(safeItems);
    } catch (err) {
      console.error("GetListData error:", err);
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

  // 🔹 Search + Status Filter Combined
  useEffect(() => {
    let data = Array.isArray(listData) ? [...listData] : [];

    // Apply status filter
    if (statusFilter !== "All") {
      data = data.filter(item => (item.Status || "").toLowerCase() === statusFilter.toLowerCase());
    }

    // (Optional) If you want to show only items "assigned to me" when status = Pending for Approval:
    // const currentName = (props.userDisplayName || "").toLowerCase();
    // if (statusFilter === "Pending for Approval") {
    //   data = data.filter(x =>
    //     ((x.RM || "").toLowerCase() === currentName) ||
    //     ((x.HOD || "").toLowerCase() === currentName)
    //   );
    // }

    // Apply search
    const q = searchTerm.trim().toLowerCase();
    if (q) {
      data = data.filter((item) => {
        const haystack = [
          item?.Title,                              // EMD Request No.
          item?.TenderNo,
          item?.VendorName?.Name,                  // Customer/Vendor Name
          item?.TenderType?.TenderType,
          String(item?.EMDAmount ?? ""),
          item?.Currency?.Currency,
          item?.Status,
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();

        return haystack.includes(q);
      });
    }

    setFilteredData(data);
    setCurrentPage(1);
  }, [searchTerm, statusFilter, listData, props.userDisplayName]);

  // 🔹 Pagination
  const totalPages = Math.ceil((filteredData?.length || 0) / itemsPerPage);

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) setCurrentPage(page);
  };

  const sortedData = useMemo(
    () => [...(filteredData || [])].sort((a, b) => (b?.Id || 0) - (a?.Id || 0)),
    [filteredData]
  );

  const paginatedData = sortedData.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  // status badge helper (optional)
  const getStatusClass = (status?: string) => {
    const s = (status || "Pending for Approval").toLowerCase();
    if (s === "pending for vouching") return "badge badge-vouching";
    if (s === "pending for payment") return "badge badge-payment";
    return "badge badge-approval"; // default: Pending for Approval
  };

  return (
    <div className="dashboard-wrapper">

      {/* Header */}
      <div className="header">
        <div className="left-banner">
          <div className="logo-text">
            <h2>Approval Dashboard</h2>
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
            onChange={(e) => setStatusFilter(e.target.value as any)}
            style={{ width: "250px" }}
          >
            <option value="All">All Status</option>
            <option value="Pending for Approval">Pending for Approval</option>
            <option value="Pending for Vouching">Pending for Vouching</option>
            <option value="Pending for Payment">Pending for Payment</option>
          </select>
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
                  <th className="px-4 py-2">View</th>
                  <th className="px-4 py-2">EMD Request No.</th>
                  <th className="px-4 py-2">EMD Requester.</th>
                  <th className="px-4 py-2">Tender No.</th>
                  <th className="px-4 py-2">Tender Type</th>
                  <th className="px-4 py-2">Customer/Vendor</th>
                  <th className="px-4 py-2">EMD Amount</th>
                  <th className="px-4 py-2">Currency</th>
                  <th className="px-4 py-2">Status</th>
                </tr>
              </thead>
              <tbody>
                {loading ? (
                  <tr>
                    <td colSpan={10} className="text-center">Loading data...</td>
                  </tr>
                ) : paginatedData.length === 0 ? (
                  <tr>
                    <td colSpan={10} className="text-center">No records found</td>
                  </tr>
                ) : (
                  paginatedData.map((item) => (
                    <tr key={item.Id}>
                      <td className="px-4 py-2">
                        <Link to={`/ViewForm?ItemId=${item.Id}`}>
                          <img src={View} width={16} alt="View" />
                        </Link>

                        <Link to={`/MANACApprovalForm?ItemId=${item.Id}`}>
                          <img src={Edit} width={16} alt="Edit" />
                        </Link>

                      </td>
                      <td className="px-4 py-2">{item.Title || "-"}</td>
                      <td className="px-4 py-2">{item.Author?.Title}</td>
                      <td className="px-4 py-2">{item.TenderType?.TenderType || "-"}</td>
                      <td className="px-4 py-2">{item.TenderNo || "-"}</td>
                      <td className="px-4 py-2">{item.VendorName?.Name || "-"}</td>
                      <td className="px-4 py-2">{item.EMDAmount ? `₹ ${formatAmount(item.EMDAmount)}` : "-"}</td>
                      <td className="px-4 py-2">{item.Currency?.Currency || "-"}</td>
                      <td className="px-4 py-2">
                        {item.Status || "Pending for Approval"}
                        {/* <span className={getStatusClass(item.Status)}>
                          {item.Status || "Pending for Approval"}
                        </span> */}
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

export default ApprovalDashboard;
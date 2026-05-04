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

type Lookup<T = string> = { Name?: T;  Title?: T; Currency?: T; TenderType?: T };

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
    if (s === "pending for payment")  return "badge badge-payment";
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

      {/* Filters */}
      <div className="filter-section">
        <div className="filter-left">
          <input
            type="text"
            placeholder="Search by EMD No, Tender No, Vendor, Amount..."
            className="form-control"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />

          {/* Status filter */}
          <select
            className="form-control"
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value as any)}
            style={{ marginLeft: 8 }}
          >
            <option value="All">All Status</option>
            <option value="Pending for Approval">Pending for Approval</option>
            <option value="Pending for Vouching">Pending for Vouching</option>
            <option value="Pending for Payment">Pending for Payment</option>
          </select>
        </div>
      </div>

      {/* Table */}
      <div className="table-section">
        <div className="table-vert-scroll">
          <table className="custom-table">
            <thead>
              <tr>
                <th>EMD Request No.</th>
                <th>EMD Requester.</th>
                <th>Tender Type</th>
                <th>Tender No.</th>
                <th>Customer/Vendor</th>
                <th>EMD Amount</th>
                <th>Currency</th>
                <th>Status</th>
                <th>Action</th>
              </tr>
            </thead>

            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={8} className="text-center">Loading data...</td>
                </tr>
              ) : paginatedData.length === 0 ? (
                <tr>
                  <td colSpan={8} className="text-center">No records found</td>
                </tr>
              ) : (
                paginatedData.map((item) => (
                  <tr key={item.Id}>
                    <td>{item.Title || "-"}</td>
                   <td>{item.Author?.Title}</td>
                    <td>{item.TenderType?.TenderType || "-"}</td>
                    <td>{item.TenderNo || "-"}</td>
                    <td>{item.VendorName?.Name || "-"}</td>
                    <td>{item.EMDAmount ? `₹ ${formatAmount(item.EMDAmount)}` : "-"}</td>
                    <td>{item.Currency?.Currency || "-"}</td>

                    <td>
                      <span className={getStatusClass(item.Status)}>
                        {item.Status || "Pending for Approval"}
                      </span>
                    </td>

                    <td>
                      {/* {item.Status === "Pending for Approval" ? (
                        <Link to={`/MANACApprovalForm?ItemId=${item.Id}`}>
                          <img src={Edit} width={16} alt="Edit" />
                        </Link>
                      ) : (
                        <Link to={`/MANACApprovalForm?ItemId=${item.Id}`}>
                          <img src={View} width={16} alt="View" />
                        </Link>
                      )} */}
                          <Link to={`/ViewForm?ItemId=${item.Id}`}>
                          <img src={View} width={16} alt="View" />
                        </Link>

                        <Link to={`/MANACApprovalForm?ItemId=${item.Id}`}>
                          <img src={Edit} width={16} alt="Edit" />
                        </Link>

                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        <div className="pagination">
          <button onClick={() => handlePageChange(currentPage - 1)} disabled={currentPage === 1}>
            <img src={Left} width={14} alt="Previous" />
          </button>

          {Array.from({ length: totalPages }, (_, i) => i + 1)
            .filter((page) => Math.abs(page - currentPage) <= 2)
            .map((page) => (
              <button
                key={page}
                onClick={() => handlePageChange(page)}
                className={currentPage === page ? "active" : ""}
              >
                {page}
              </button>
            ))}

          <button
            onClick={() => handlePageChange(currentPage + 1)}
            disabled={currentPage === totalPages || totalPages === 0}
          >
            <img src={Right} width={14} alt="Next" />
          </button>
        </div>
      </div>
    </div>
  );
};

export default ApprovalDashboard;
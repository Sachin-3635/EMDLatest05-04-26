import * as React from "react";
import { useState, useEffect, useMemo } from "react";
import { Link, useHistory } from "react-router-dom";
import type { ISonaEmdProps } from "../ISonaEmdProps";
import "../../components/Pages/Css/InitiatorDashboard.scss";
import SPCRUDOPS from "../../service/BAL/spcrud";

import Left from "../../assets/LeftArrow.png";
import Right from "../../assets/RightArrow.png";
import View from "../../assets/Eye.png";
import Edit from "../../assets/Pencil.png";

type Lookup<T = string> = { Name?: T;  Title?: T; Currency?: T; TenderType?: T };

interface EmdItem {
  Id: number;
  Title?: string;                 // EMD Request No. (Title)
  TenderNo?: string;
  EMDAmount?: string | number;
  Status?: string;
  Author?: Lookup<string>;        // Requester (Author)

  // Lookups
  VendorName?: Lookup<string>;
  TenderType?: Lookup<string>;
  Currency?: Lookup<string>;

  Created?: string;
  Modified?: string;
}

export const TreasuryDashboard: React.FC<ISonaEmdProps> = (props: ISonaEmdProps) => {
  const [searchTerm, setSearchTerm] = useState("");
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
      .toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" })
      .replace(/\//g, "-");
  };

  const formatAmount = (value?: string | number) => {
    if (value === undefined || value === null || value === "") return "-";
    const num = typeof value === "number" ? value : Number(value);
    if (isNaN(num)) return String(value);
    return num.toLocaleString("en-IN");
  };

  // Optional status badge helper
  const getStatusClass = (status?: string) => {
    const s = (status || "").toLowerCase();
    if (s === "pending for payment") return "badge badge-payment";
    if (s === "pending for vouching") return "badge badge-vouching";
    return "badge badge-approval";
  };

  // 🔹 Fetch Data: ONLY Pending for Payment
  const GetListData = async () => {
    try {
      setLoading(true);
      const spCrudOps = await SPCRUDOPS();

      const selectFields = [
        "Id",
        "Title",                // EMD Request No.
        "Author/Title",
        "TenderNo",
        "EMDAmount",
        "Status",               // Choice
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
const currentUserId = props.context.pageContext.legacyPageContext.userId;
      // ✅ Treasury: only items pending for payment
      const filter = `Status eq 'Pending for Payment' and CurrentApproverId eq ${currentUserId}`;

      const items: EmdItem[] = await spCrudOps.getData(
        "EMDDetails",
        selectFields,
        expandFields,
        filter,
        { column: "Id", isAscending: false }, // latest first
        5000,
        props
      );

      const safeItems = Array.isArray(items) ? items : [];
      setListData(safeItems);
      setFilteredData(safeItems);
    } catch (err) {
      console.error("Treasury GetListData error:", err);
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

  // 🔹 Search
  useEffect(() => {
    let data = Array.isArray(listData) ? [...listData] : [];
    const q = searchTerm.trim().toLowerCase();

    if (q) {
      data = data.filter((item) => {
        const haystack = [
          item?.Title,                         // EMD Request No
          item?.TenderNo,
          item?.VendorName?.Title,
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
  }, [searchTerm, listData]);

  // 🔹 Pagination
  const totalPages = Math.ceil((filteredData?.length || 0) / itemsPerPage);

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
    }
  };

  const sortedData = useMemo(
    () => [...(filteredData || [])].sort((a, b) => (b?.Id || 0) - (a?.Id || 0)),
    [filteredData]
  );

  const paginatedData = sortedData.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  return (
    <div className="dashboard-wrapper">
      {/* Header */}
      <div className="header">
        <div className="left-banner">
          <div className="logo-text">
            <h2>Treasury Dashboard </h2>
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
          {/* Status filter not required; dashboard is dedicated to Pending for Payment */}
        </div>
      </div>

      {/* Table */}
      <div className="table-section">
        <div className="table-vert-scroll">
          <table className="custom-table">
            <thead>
              <tr>
                <th>EMD Request No.</th>
                <th>EMD Requester</th>
                <th>Tender Type</th>
                <th>Tender No.</th>
                <th>Customer/Vendor</th>
                <th>EMD Amount</th>
                <th>Currency</th>
                <th>Status</th>
                <th>View</th>
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
                    <td>{item.Author?.Title || "-"}</td>
                    <td>{item.TenderType?.TenderType || "-"}</td>
                    <td>{item.TenderNo || "-"}</td>
                    <td>{item.VendorName?.Name || "-"}</td>
                    <td>{item.EMDAmount ? `₹ ${formatAmount(item.EMDAmount)}` : "-"}</td>
                    <td>{item.Currency?.Currency || "-"}</td>
                    <td>
                      <span className={getStatusClass(item.Status)}>
                        {item.Status || "-"}
                      </span>
                    </td>
                    <td>
                      {/* 👉 View or Payment action form:
                          Change to your payment page if exists:
                          <Link to={`/TreasuryPaymentForm?ItemId=${item.Id}`}> ... </Link>
                      */}
                      <Link to={`/ViewForm?ItemId=${item.Id}`}>
                        <img src={View} width={16} alt="View" />
                      </Link>
                      <Link to={`/UTRDetailEntry?ItemId=${item.Id}`}>
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

export default TreasuryDashboard;

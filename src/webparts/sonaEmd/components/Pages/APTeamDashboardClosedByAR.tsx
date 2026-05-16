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

type Lookup<T = string> = { Name?: T; Title?: T; Currency?: T; TenderType?: T };

interface EmdItem {
  Id: number;
  Title?: string;
  TenderNo?: string;
  EMDAmount?: string | number;
  Status?: string;

  // Lookups
  VendorName?: Lookup<string>;
  TenderType?: Lookup<string>;
  Currency?: Lookup<string>;

  Created?: string;
  Modified?: string;
}

export const APTeamDashboardClosedByAR: React.FC<ISonaEmdProps> = (props: ISonaEmdProps) => {
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

  // Status badge helper (optional styling)
  const getStatusClass = (status?: string) => {
    const s = (status || "").toLowerCase().trim();
    if (s === "pending for closure vouching") return "badge badge-pending-closure-vouching";
    if (s === "pending for vouching") return "badge badge-vouching";
    if (s === "pending for payment") return "badge badge-payment";
    if (s === "closed") return "badge badge-closed";
    if (s === "rejected") return "badge badge-rejected";
    if (s === "send back") return "badge badge-sendback";
    return "badge badge-approval"; // default fallback
  };

  // 🔹 Fetch Data: ONLY "Pending for Closure Vouching"
  const GetListData = async () => {
    try {
      setLoading(true);
      const spCrudOps = await SPCRUDOPS();

      const selectFields = [
        "Id",
        "Title",
        "TenderNo",
        "EMDAmount",
        "Status", // Choice
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

      const expandFields = "VendorName,TenderType,Currency";

      // ✅ Fetch only items "Pending for Closure Vouching" (ensure exact choice text)
      // const filter = "Status eq 'Pending for Closure Vouching'";
      const currentUserId = props.context.pageContext.legacyPageContext.userId;

      const filter = `Status eq 'Pending for Closure Vouching' and CurrentApproverId eq ${currentUserId}`;

      const items: EmdItem[] = await spCrudOps.getData(
        "EMDDetails",
        selectFields,
        expandFields,
        filter,
        { column: "Id", isAscending: false },
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

  // 🔹 Search
  useEffect(() => {
    let data = Array.isArray(listData) ? [...listData] : [];
    const q = searchTerm.trim().toLowerCase();

    if (q) {
      data = data.filter((item) => {
        const haystack = [
          item?.Title,
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
            <h2>AP Team Dashboard for Closure Vouching</h2>
          </div>
        </div>
      </div>

      <div className="col-md-12 d-flex  px-2 py-2 justify-content-between align-items-center" style={{ gap: "5px" }}>
        <input type="text" placeholder="Search by Tender No, Vendor, Amount..."
          className="form-control" style={{ width: "250px" }}
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
        />
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
                  <th className="px-4 py-2">Tender Type</th>
                  <th className="px-4 py-2">Tender No.</th>
                  <th className="px-4 py-2">Customer/Vendor</th>
                  <th className="px-4 py-2">EMD Amount</th>
                  <th className="px-4 py-2">Currency</th>
                  <th className="px-4 py-2">Status</th>
                  <th className="px-4 py-2">Action</th>
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
                      <td className="px-4 py-2">{item.Title || "-"}</td>
                      <td className="px-4 py-2">{item.TenderType?.TenderType || "-"}</td>
                      <td className="px-4 py-2">{item.TenderNo || "-"}</td>
                      <td className="px-4 py-2">{item.VendorName?.Name || "-"}</td>
                      <td className="px-4 py-2">{item.EMDAmount ? `₹ ${formatAmount(item.EMDAmount)}` : "-"}</td>
                      <td className="px-4 py-2">{item.Currency?.Currency || "-"}</td>
                      <td className="px-4 py-2">
                        {item.Status || "-"}
                        {/* <span className={getStatusClass(item.Status)}>
                          {item.Status || "-"}
                        </span> */}
                      </td>
                      <td className="px-4 py-2">
                        {/* Route ko aap apne AP/Approval flow ke mutabik change kar sakte ho */}
                        {/* Example: AR Closure Approval Form (read-only/approval) */}
                        <Link to={{ pathname: "/EMDClosureRequestForm", search: `?ItemId=${item.Id}` }}>
                          <img src={View} width={16} alt="View" />
                        </Link>
                        <Link to={{ pathname: "/EMDClosureRequestForm", search: `?ItemId=${item.Id}` }}>
                          <img src={Edit} width={16} alt="Edit" />
                        </Link>

                        {/* Example: If AP needs to perform vouching here, point to AP vouching form */}
                        {/* <Link to={{ pathname: "/VouchingbyAPTeamForm", search: `?ItemId=${item.Id}` }}>
                        <img src={Edit} width={16} alt="Edit" />
                      </Link> */}
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

export default APTeamDashboardClosedByAR;

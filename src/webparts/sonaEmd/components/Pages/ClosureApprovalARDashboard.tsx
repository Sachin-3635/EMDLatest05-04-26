import * as React from "react";
import { useState, useEffect } from "react";
import { Link } from "react-router-dom";
import type { ISonaEmdProps } from "../ISonaEmdProps";
import "../../components/Pages/Css/InitiatorDashboard.scss";
import SPCRUDOPS from "../../service/BAL/spcrud";

import Left from "../../assets/LeftArrow.png";
import Right from "../../assets/RightArrow.png";
import View from "../../assets/Eye.png";
import Edit from "../../assets/Pencil.png";

export const ClosureApprovalARDashboard: React.FC<ISonaEmdProps> = (props: ISonaEmdProps) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [listData, setListData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  // Pagination
  const itemsPerPage = 10;
  const [currentPage, setCurrentPage] = useState(1);

  const formatDate = (dateString: string) => {
    if (!dateString) return "";
    const date = new Date(dateString);
    return date
      .toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" })
      .replace(/\//g, "-");
  };

  const toNumeric = (val: any) => {
    if (val == null) return 0;
    const s = String(val).replace(/,/g, "");
    const n = Number(s);
    return Number.isNaN(n) ? 0 : n;
  };

  // 🔹 Fetch "Pending for Closure Approval" items from EMDDetails
  const GetListData = async () => {
    setLoading(true);

    const spCrudOps = await SPCRUDOPS();
    const listName = "EMDDetails";

    const select = [
      "ID",
      "Title",
      "EmployeeCode",
      "EmployeeName",
      "Email",
      "ContactNo",
      "Division",
      "Location",
      "RM",
      "HOD",
      "EmployeeStatus",
      "VendorName/Id",
      "VendorName/Title",
      "VendorName/Name",
      "VendorCode",
      "VendorSite",
      "TenderType/Id",
      "TenderType/Title",
      "TenderNo",
      "TenderDate",
      "TenderAmount",
      "ContractType/Id",
      "ContractType/Title",
      "ProductType/Id",
      "ProductType/Title",
      "EMDAmount",
      "EMDPercentage",
      "Currency/Id",
      "Currency/Currency",
      "TenderClosingDate",
      "ModeofPayment/Id",
      "ModeofPayment/Title",
      "Status",
      "ApproverComment",
      "GLCode",
      "VoucherNo",
      "VouchingDate",
      "APTeamComment",
      "UTRNo",
      "UTRDate",
      "TreasuryComment",
      "Modified",
      "Created",
      "Author/Id",
      "Author/Title",
      "Editor/Id",
      "Editor/Title",
    ].join(",");

    const expand = [
      "VendorName",
      "TenderType",
      "ContractType",
      "ProductType",
      "Currency",
      "ModeofPayment",
      "Author",
      "Editor",
    ].join(",");

    // ✅ Only "Pending for Closure Approval"
   // const filter = `Status eq 'Pending for Closure Approval'`;

    const currentUserId = props.context.pageContext.legacyPageContext.userId;

const filter = `Status eq 'Pending for Closure Approval' and CurrentApproverId eq ${currentUserId}`;

    const parentItems = await spCrudOps.getRootData(
      listName,
      select,
      expand,
      filter,
      { column: "ID", isAscending: false },
      5000,
      props
    );

    const items = parentItems || [];
    setListData(items);
    setFilteredData(items);
    setLoading(false);
  };

  useEffect(() => {
    GetListData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 🔹 Search (within "Pending for Closure Approval" results)
  useEffect(() => {
    let data = [...listData];

    if (searchTerm) {
      const lowerSearch = searchTerm.toLowerCase();

      data = data.filter((item) => {
        const haystack = [
          item.Title,
          item.EmployeeCode,
          item.EmployeeName,
          item.Email,
          item.ContactNo,
          item.Division,
          item.Location,
          item.RM,
          item.HOD,
          item.EmployeeStatus,
          item.VendorName?.Title,
          item.VendorCode,
          item.VendorSite,
          item.TenderType?.Title,
          item.TenderNo,
          item.TenderDate ? formatDate(item.TenderDate) : "",
          item.TenderAmount,
          item.ContractType?.Title,
          item.ProductType?.Title,
          item.EMDAmount,
          item.EMDPercentage,
          item.Currency?.Title,
          item.TenderClosingDate ? formatDate(item.TenderClosingDate) : "",
          item.ModeofPayment?.Title,
          item.Status,
          item.ApproverComment,
          item.GLCode,
          item.VoucherNo,
          item.VouchingDate ? formatDate(item.VouchingDate) : "",
          item.APTeamComment,
          item.UTRNo,
          item.UTRDate ? formatDate(item.UTRDate) : "",
          item.TreasuryComment,
        ]
          .map((v) => (v == null ? "" : String(v)))
          .join(" ")
          .toLowerCase();

        return haystack.includes(lowerSearch);
      });
    }

    setFilteredData(data);
    setCurrentPage(1);
  }, [searchTerm, listData]);

  // 🔹 Pagination
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
    }
  };

  const sortedData = [...filteredData].sort((a, b) => b.ID - a.ID);

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
            <h2>Closure Approval AR Dashboard</h2>
            {/* <div className="subtitle">Showing: Pending for Closure Approval</div> */}
          </div>
        </div>
      </div>

      {/* Filters (only search; status is locked) */}
      <div className="filter-section">
        <div className="filter-left">
          <input
            type="text"
            placeholder="Search..."
            className="form-control"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
          {/* Status dropdown removed intentionally; locked to 'Pending for Closure Approval' */}
        </div>
      </div>

      {/* Table */}
      <div className="table-section">
        <div className="table-vert-scroll">
          <table className="custom-table">
            <thead>
              <tr>
                <th>EMD Request No.</th>
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
                  <td colSpan={8} className="text-center">
                    Loading data...
                  </td>
                </tr>
              ) : paginatedData.length === 0 ? (
                <tr>
                  <td colSpan={8} className="text-center">
                    No records found
                  </td>
                </tr>
              ) : (
                paginatedData.map((item) => (
                  <tr key={item.ID}>
                    {/* If no dedicated EMDRequestNo, fallback to Title/TenderNo/ID */}
                    <td>{item.Title || item.TenderNo || item.ID}</td>
                    <td>{item.TenderType?.Title || "-"}</td>
                    <td>{item.TenderNo || "-"}</td>
                    <td>{item.VendorName?.Name || "-"}</td>
                    <td>
                      {item.EMDAmount
                        ? toNumeric(item.EMDAmount).toLocaleString("en-IN")
                        : "-"}
                    </td>
                    <td>{item.Currency?.Currency || "-"}</td>
                    <td>
                      <span className={`status-badge ${item.Status?.replace(/\s+/g, "-")}`}>
                        {item.Status || "-"}
                      </span>
                    </td>
                    <td>
                      {/* Navigate to Closure Request Form using ItemId (same as your other dashboard) */}
                      <Link
                        to={{
                          pathname: "/ViewForm",
                          search: `?ItemId=${item.ID}`,
                        }}
                      >
                        <img src={View} width={16} alt="View" />
                      </Link>
                      <Link
                        to={{
                          pathname: "/ARClosureApprovalForm",
                          search: `?ItemId=${item.ID}`,
                        }}
                      >
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
          <button
            onClick={() => handlePageChange(currentPage - 1)}
            disabled={currentPage === 1}
          >
            <img src={Left} width={14} />
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
            <img src={Right} width={14} />
          </button>
        </div>
      </div>
    </div>
  );
};

export default ClosureApprovalARDashboard;
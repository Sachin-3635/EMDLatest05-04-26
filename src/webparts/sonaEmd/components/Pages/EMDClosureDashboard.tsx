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

export const EMDClosureDashboard: React.FC<ISonaEmdProps> = (props: ISonaEmdProps) => {
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
      .toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
      })
      .replace(/\//g, "-");
  };

  const toNumeric = (val: any) => {
    if (val == null) return 0;
    const s = String(val).replace(/,/g, "");
    const n = Number(s);
    return Number.isNaN(n) ? 0 : n;
  };

  // 🔹 Fetch Data from EMDDetails (ONLY EMD Paid)
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
      "TenderType/TenderType",
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

    // ✅ Only EMD Paid
    //    const filter = `Status eq 'EMD Paid'`; and CurrentApproverId eq ${currentUserId}
    const currentUserId = props.context.pageContext.legacyPageContext.userId;

    const filter = `Status eq 'EMD Paid'`;

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

  // 🔹 Search (within EMD Paid results)
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
          item.TenderType?.TenderType,
          item.TenderNo,
          item.TenderDate ? formatDate(item.TenderDate) : "",
          item.TenderAmount,
          item.ContractType?.Title,
          item.ProductType?.Title,
          item.EMDAmount,
          item.EMDPercentage,
          item.Currency?.Currency,
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
            <h2>EMD Closure Dashboard</h2>
          </div>
        </div>
      </div>

      <div className='col-md-12 px-2 py-2 d-flex justify-content-between align-items-center flex-wrap'>
        <input
          type="text"
          placeholder="Search..."
          className="form-control"
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          style={{ width: "250px" }}
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
                  <th className="px-4 py-2">View</th>
                  <th className="px-4 py-2">EMD Request No.</th>
                  <th className="px-4 py-2">Tender Type</th>
                  <th className="px-4 py-2">Tender No.</th>
                  <th className="px-4 py-2">Customer/Vendor</th>
                  <th className="px-4 py-2">EMD Amount</th>
                  <th className="px-4 py-2">Currency</th>
                  <th className="px-4 py-2">Status</th>
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
                      <td className="px-4 py-2">
                        {/*  Nvaigate to form */}
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
                            pathname: "/ClosureRequestForm",
                            search: `?ItemId=${item.ID}`,
                          }}
                        >
                          <img src={Edit} width={16} alt="Edit" />
                        </Link>
                      </td>
                      {/* Using Title or TenderNo or fallback to ID */}
                      <td className="px-4 py-2">{item.Title || item.TenderNo || item.ID}</td>
                      <td className="px-4 py-2">{item?.TenderType?.TenderType ?? "-"}</td>
                      {/* <td className="px-4 py-2">{item.TenderType?.Title || "-"}</td> */}
                      <td className="px-4 py-2">{item.TenderNo || "-"}</td>
                      <td className="px-4 py-2">{item.VendorName?.Name || "-"}</td>
                      <td className="px-4 py-2">
                        {item.EMDAmount
                          ? toNumeric(item.EMDAmount).toLocaleString("en-IN")
                          : "-"}
                      </td>
                      <td className="px-4 py-2">{item.Currency?.Currency || "-"}</td>
                      <td className="px-4 py-2">
                        <span
                          className={`status-badge ${item.Status?.replace(/\s+/g, "-")}`}
                        >
                          {item.Status || "-"}
                        </span>
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

export default EMDClosureDashboard;
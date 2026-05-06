import * as React from "react";
import { useState, useEffect } from "react";
import { useHistory, Link } from "react-router-dom";
import type { ISonaEmdProps } from "../ISonaEmdProps";
import "../../components/Pages/Css/InitiatorDashboard.scss";
import SPCRUDOPS from "../../service/BAL/spcrud";

import Left from "../../assets/LeftArrow.png";
import Right from "../../assets/RightArrow.png";
import View from "../../assets/Eye.png";
import Edit from "../../assets/Pencil.png";
import Renew from "../../assets/Renew.png";

export const InitiatorDashboard: React.FC<ISonaEmdProps> = (
  props: ISonaEmdProps
) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");
  const [listData, setListData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  // Pagination
  const itemsPerPage = 10;
  const [currentPage, setCurrentPage] = useState(1);

  const history = useHistory();

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

  // 🔹 Fetch Data
  const GetListData = async () => {
    setLoading(true);

    const spCrudOps = await SPCRUDOPS();

    const parentItems = await spCrudOps.getRootData(
      "ForexRequest",
      "ID,ForexNumber,EmployeeName,EmployeeCode,RequestedOn,Status,Location,VendorName,TotalAmount,Author/Id",
      "Author",
      "AuthorId eq " + props.id,
      { column: "ID", isAscending: false },
      5000,
      props
    );

    setListData(parentItems);
    setFilteredData(parentItems);
    setLoading(false);
  };

  useEffect(() => {
    GetListData();
  }, []);

  // 🔹 Search + Status Filter Combined
  useEffect(() => {
    let data = [...listData];

    if (statusFilter !== "All") {
      data = data.filter((item) => item.Status === statusFilter);
    }

    if (searchTerm) {
      const lowerSearch = searchTerm.toLowerCase();

      data = data.filter((item) =>
        Object.values({
          requestNo: item.ForexNumber,
          // requestDate: formatDate(item.RequestDate),
          vendorName: item.VendorName,
          EmployeeCode: item.EmployeeCode,
          EmployeeName: item.EmployeeName,
          requestDate: formatDate(item.RequestedOn),
          status: item.Status,
          amount: item.TotalAmount,
        })
          .join(" ")
          .toLowerCase()
          .includes(lowerSearch)
      );
    }

    setFilteredData(data);
    setCurrentPage(1);
  }, [searchTerm, statusFilter, listData]);

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

      {/* Header (UNCHANGED) */}
      <div className="header">
        <div className="left-banner">
          <div className="logo-text">
            <h2>Forex Initiation Dashboard</h2>
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
            <option value="All">All Status</option>
            <option value="Draft">Draft</option>
            <option value="Submitted">Submitted</option>
            <option value="Approved">Approved</option>
            <option value="Send back">Send back</option>
            <option value="Rejected">Rejected</option>
          </select>
        </div>
        <div className='Dashbaordcreatebutton'>
          <Link to="/EMDRequestForm" className="create-button"> + New Request </Link>
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
                  <th className="px-4 py-2">Request No.</th>
                  <th className="px-4 py-2">EmployeeCode</th>
                  <th className="px-4 py-2">Employee Name</th>
                  <th className="px-4 py-2">Request Date</th>
                  <th className="px-4 py-2">Location</th>
                  <th className="px-4 py-2">Vendor Name</th>
                  <th className="px-4 py-2">Amount</th>
                  <th className="px-4 py-2">Status</th>
                  <th className="px-4 py-2">Action</th>
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
                      <td className="px-4 py-2">{item.ForexNumber}</td>
                      <td className="px-4 py-2">{item.EmployeeCode}</td>
                      <td className="px-4 py-2">{item.EmployeeName}</td>
                      <td className="px-4 py-2">{formatDate(item.RequestedOn)}</td>
                      <td className="px-4 py-2">{item.Location}</td>
                      <td className="px-4 py-2">{item.VendorName}</td>
                      <td className="px-4 py-2">₹ {item.TotalAmount || "-"}</td>
                      <td className="px-4 py-2">
                        <span
                          className={`status-badge ${item.Status?.replace(" ", "-")}`}
                        >
                          {item.Status}
                        </span>
                      </td>
                      <td className="px-4 py-2">
                        {item.Status === "Send back" ||
                          item.Status === "Draft" ? (
                          <Link to={`/ViewForm/${item.ID}`}>
                            <img src={Edit} width={16} alt="Edit" />
                          </Link>
                        ) : (
                          <>
                            <Link to={`/EditRequest/${item.ID}`}>
                              <img src={View} width={16} alt="View" />
                            </Link>


                          </>
                        )}
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

export default InitiatorDashboard;

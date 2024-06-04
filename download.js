// download excel=====================================================================================================================================================================

const handleDownloadExcel = () => {
    // Define the columns for Excel
    const columns = [
      "Sr No.",
      "Hotel ID",
      "Hotel Name",
      "Hotel Address",
      "State",
      "District",
      "City",
      "Pincode",
      "Contact No.",
      "Email",
      "Owner First Name",
      "Owner Last Name",
      "Owner Address",
      "Owner State",
      "Owner District",
      "Owner City",
      "Owner Pincode",
      "Guest Name",
      "Age",
      "Gender",
      "Country",
      "State",
      "District",
      "Pincode",
      "Contact No.",
      "Document Type",
      "Document ID",
      "Document Image",
      "Check-in Date/Time",
      "Check-out Date/Time",
      "Room No.",
      "No. of Persons",
      "No. of Adults",
      "No. of Children",
      "Coming From",
      "Going To",
    ];

    // Extract data for Excel
    const data = filteredHotels.map((row, index) => [
      index + 1,
      row.hotel.hotel_user_name,
      row.hotel.hotel_name,
      row.hotel.hotel_address,
      row.hotel.state,
      row.hotel.district,
      row.hotel.city,
      row.hotel.pincode,
      row.hotel.mobile,
      row.hotel.email,
      row.hotel.owner_first_name,
      row.hotel.owner_last_name,
      row.hotel.owner_house_address,
      row.hotel.owner_state,
      row.hotel.owner_district,
      row.hotel.owner_city,
      row.hotel.owner_pincode,
      row.guest_name,
      row.age,
      row.gender,
      row.country,
      row.state,
      row.district,
      row.pincode,
      row.mobile,
      row.document_type,
      row.document_no_field,
      "document image",
      row.check_in,
      row.check_out,
      row.room_no_field,
      row.no_person,
      row.adults,
      row.child,
      row.coming_from,
      row.going_to,
    ]);

    // Create a new Excel workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([columns, ...data]);
    XLSX.utils.book_append_sheet(wb, ws, "Guests List");

    // Apply basic styling to the header row
    const headerCellStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "FF0000" } }, // Header background color
    };
    XLSX.utils.sheet_add_json(ws, [], { header: columns, origin: "A1" });
    for (let i = 0; i < columns.length; i++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: i });
      ws[cellAddress].s = headerCellStyle;
    }

    // Add data to the sheet
    XLSX.utils.sheet_add_json(ws, data, { skipHeader: true, origin: "A2" });

    // Save Excel file
    XLSX.writeFile(wb, "guests_list_report.xlsx");
  };

// download pdf=====================================================================================================================================================================

const handleDownloadPDF = () => {
    console.log(filteredHotels);
    const doc = new jsPDF({
      orientation: "landscape", // Set landscape orientation
    });

    const tableData = filteredHotels.map((row, index) => [
      index + 1,
      row.hotel.hotel_user_name,
      row.hotel.hotel_name,
      row.hotel.hotel_address,
      row.hotel.state,
      row.hotel.district,
      row.hotel.city,
      row.hotel.pincode,
      row.hotel.mobile,
      row.hotel.email,
      row.hotel.owner_first_name,
      row.hotel.owner_last_name,
      row.hotel.owner_house_address,
      row.hotel.owner_state,
      row.hotel.owner_district,
      row.hotel.owner_city,
      row.hotel.owner_pincode,
      row.guest_name,
      row.age,
      row.gender,
      row.country,
      row.state,
      row.district,
      row.pincode,
      row.mobile,
      row.document_type,
      row.document_no_field,
      "document image",
      row.check_in,
      row.check_out,
      row.room_no_field,
      row.no_person,
      row.adults,
      row.child,
      row.coming_from,
      row.going_to,
    ]);

    const tableStyles = {
      startY: 20, // Start position Y
      margin: { top: 10, left: 10, right: 10, bottom: 10 }, // Margin
      styles: { cellPadding: 1, fontSize: 4, valign: "middle" }, // Cell styles
      headStyles: { fillColor: [100, 100, 100] }, // Header styles
      columnStyles: { 0: { cellWidth: "auto" } }, // Set first column width to auto
    };

    doc.autoTable({
      head: [[
        "Sr No.",
        "Hotel ID",
        "Hotel Name",
        "Hotel Address",
        "State",
        "District",
        "City",
        "Pincode",
        "Contact No.",
        "Email",
        "Owner First Name",
        "Owner Last Name",
        "Owner Address",
        "Owner State",
        "Owner District",
        "Owner City",
        "Owner Pincode",
        "Guest Name",
        "Age",
        "Gender",
        "Country",
        "State",
        "District",
        "Pincode",
        "Contact No.",
        "Document Type",
        "Document ID",
        "Document Image",
        "Check-in Date/Time",
        "Check-out Date/Time",
        "Room No.",
        "No. of Persons",
        "No. of Adults",
        "No. of Children",
        "Coming From",
        "Going To",
      ]],
      body: tableData,
      ...tableStyles,
    });

  doc.save("guests_list_report.pdf");
};
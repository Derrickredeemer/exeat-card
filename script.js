const { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun } = docx;

function downloadWord() {
  // Create document
  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({ text: "Selula Royal School of Excellence", heading: "Heading1" }),
        new Paragraph({ text: "Boarders Timetable", heading: "Heading2" }),
        new Paragraph({ text: "Weekdays" }),
        createTable([
          ["Day", "Morning", "Evening"],
          ["Monday", "Maths", "Dangme"],
          ["Tuesday", "R.M.E", "English"],
          ["Wednesday", "Science", "C.A.D"],
          ["Thursday", "French", "Social"],
          ["Friday", "Maths", "English"],
        ]),
        new Paragraph({ text: "Weekends" }),
        createTable([
          ["Day", "Morning", "Evening"],
          ["Saturday", "Science (3am - 5am)", "Social (6pm - 8pm)"],
          ["Sunday", "C. Tech (3am - 5am)", "Computing (7pm - 9pm)"],
        ])
      ]
    }]
  });

  Packer.toBlob(doc).then(blob => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Selula_Boarders_Timetable.docx";
    a.click();
    URL.revokeObjectURL(url);
  });
}

function createTable(data) {
  return new Table({
    rows: data.map(row => new TableRow({
      children: row.map(cell => new TableCell({
        children: [new Paragraph({ text: cell })]
      }))
    }))
  });
}

function downloadPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF("p", "pt", "a4");
  const timetable = document.getElementById("timetable").innerText;
  doc.setFontSize(12);
  const lines = doc.splitTextToSize(timetable, 500);
  doc.text(lines, 40, 40);
  doc.save("Selula_Boarders_Timetable.pdf");
}

const readline = require("readline");
const fs = require("fs");
const Table = require("cli-table3");
const {
  Document, Packer, Table: DocxTable, TableRow, TableCell, Paragraph, TextRun,
  WidthType, AlignmentType, HeadingLevel, ShadingType, BorderStyle
} = require("docx");

const names = [
  "Hemant Vitthalrao Jatal", "Janhvi Ravindra Bhandarkar", "Kruti Milind Bagwe",
  "Diksha Manoj Patil", "Dhruvi Pradeepkumar Patel", "Riddhi Nilesh Patil",
  "Vaishnavi Mahesh Koyande", "Shreeya Mahesh Bhalerao", "Sanskruti Subhash Kadam",
  "Adarsh Arvind Gupta", "Juan Allwin Andrades", "P. S. PRIYADARSHINI",
  "Priyanka Haridas Belkhede", "Piyush Rajesh Parate", "Atharva Atul Kawtikwar",
  "Spandan Deb", "Jitesh Dhanraj Kotian", "Kaustubh Kiran Gaikwad",
  "Riyaz Ismail Memon", "Anish Ravi Nadar", "Dhanashree Sanjay Chila",
  "Vrushabh Kailas Dhomse", "Sumit Jagdish Desai", "Vaishav Rajendraprasad Pal",
  "Jai Rajesh Jain", "Luvv Sandesh Swami", "Rohan Patil", "Gracy Mitesh Shah",
  "Ruchit Ramesh Shetty", "Vanshika Rajkumar Raheja", "Ruth Anand Bonala",
  "Bhoomi Uttam Barbole", "Aditya Nagesh Raorane", "Nishant Sachin Khetal"
];

function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function getSizes(total, teamCount) {
  const base = Math.floor(total / teamCount);
  const remainder = total % teamCount;
  const sizes = [];
  for (let i = 0; i < teamCount; i++) sizes.push(base + (i < remainder ? 1 : 0));
  return sizes;
}

function buildTeams(names, teamCount) {
  const shuffled = shuffle(names);
  const sizes = getSizes(shuffled.length, teamCount);
  const teams = [];
  let idx = 0;
  for (const size of sizes) {
    teams.push(shuffled.slice(idx, idx + size));
    idx += size;
  }
  return teams;
}

// ---------- Terminal table (compact width, full horizontal lines between every name) ----------
function wrapToTwoLines(name) {
  const words = name.split(" ");
  if (words.length <= 1) return name;
  let bestSplit = 1, bestDiff = Infinity;
  for (let i = 1; i < words.length; i++) {
    const line1 = words.slice(0, i).join(" ");
    const line2 = words.slice(i).join(" ");
    const diff = Math.abs(line1.length - line2.length);
    if (diff < bestDiff) { bestDiff = diff; bestSplit = i; }
  }
  return words.slice(0, bestSplit).join(" ") + "\n" + words.slice(bestSplit).join(" ");
}

function printTerminalTable(teams) {
  const headers = teams.map((_, i) => `Team ${i + 1}`);
  const rowCount = Math.max(...teams.map(t => t.length));
  const COL_WIDTH = 18;

  const table = new Table({
    head: headers,
    style: { head: [], border: [] }, // no "compact" -> horizontal line drawn between every row
    colWidths: teams.map(() => COL_WIDTH),
    wordWrap: true
  });

  for (let r = 0; r < rowCount; r++) {
    table.push(teams.map(team => (team[r] ? wrapToTwoLines(team[r]) : "")));
  }

  console.log("\n" + table.toString() + "\n");
  console.log("Team sizes:", teams.map(t => t.length).join(", "));
}

// ---------- Word document ----------
function generateDocx(teams, teamCount, outputPath) {
  const rowCount = Math.max(...teams.map(t => t.length));
  const TABLE_WIDTH = 9360; // US Letter, 1" margins
  const COL_WIDTH = Math.floor(TABLE_WIDTH / teamCount);
  const colWidths = teams.map(() => COL_WIDTH);

  const cellBorder = {
    top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  };

  function makeCell(text, { header = false, width } = {}) {
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      borders: cellBorder,
      shading: header ? { type: ShadingType.CLEAR, fill: "D9D9D9" } : undefined,
      margins: { top: 100, bottom: 100, left: 100, right: 100 },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text, bold: header, size: 20 })]
        })
      ]
    });
  }

  const headerRow = new TableRow({
    tableHeader: true,
    children: teams.map((_, i) => makeCell(`Team ${i + 1}`, { header: true, width: colWidths[i] }))
  });

  const bodyRows = [];
  for (let r = 0; r < rowCount; r++) {
    bodyRows.push(new TableRow({
      children: teams.map((team, i) => makeCell(team[r] || "", { width: colWidths[i] }))
    }));
  }

  const table = new DocxTable({
    width: { size: TABLE_WIDTH, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...bodyRows]
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
        }
      },
      children: [
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `Team Allocation (${teamCount} Teams)` })]
        }),
        new Paragraph({ text: "" }),
        table,
        new Paragraph({ text: "" }),
        new Paragraph({
          children: [new TextRun({ text: `Team sizes: ${teams.map(t => t.length).join(", ")}`, italics: true, size: 18 })]
        })
      ]
    }]
  });

  return Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(outputPath, buffer);
  });
}

// ---------- Run ----------
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

rl.question("How many teams do you want? (Enter 4 or 5): ", async (answer) => {
  const teamCount = parseInt(answer.trim(), 10);

  if (teamCount !== 4 && teamCount !== 5) {
    console.log("Invalid input. Please enter 4 or 5.");
    rl.close();
    return;
  }

  const teams = buildTeams(names, teamCount);

  printTerminalTable(teams);

  const outputPath = `Teams_${teamCount}.docx`;
  await generateDocx(teams, teamCount, outputPath);
  console.log(`Word document saved as: ${outputPath}`);

  rl.close();
});
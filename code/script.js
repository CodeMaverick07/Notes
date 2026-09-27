const readline = require("readline");
const Table = require("cli-table3");

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
  for (let i = 0; i < teamCount; i++) {
    sizes.push(base + (i < remainder ? 1 : 0));
  }
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

function printTeamsTable(teams) {
  const headers = teams.map((_, i) => `Team ${i + 1}`);
  const rowCount = Math.max(...teams.map(t => t.length));

  const table = new Table({
    head: headers,
    style: { head: [], border: [] },
    colWidths: teams.map(t => Math.max(...t.map(n => n.length), 10) + 4)
  });

  for (let r = 0; r < rowCount; r++) {
    table.push(teams.map(team => team[r] || ""));
  }

  console.log("\n" + table.toString() + "\n");
  console.log("Team sizes:", teams.map(t => t.length).join(", "));
}

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

rl.question("How many teams do you want? (Enter 4 or 5): ", (answer) => {
  const teamCount = parseInt(answer.trim(), 10);

  if (teamCount !== 4 && teamCount !== 5) {
    console.log("Invalid input. Please enter 4 or 5.");
    rl.close();
    return;
  }

  const teams = buildTeams(names, teamCount);
  printTeamsTable(teams);
  rl.close();
});

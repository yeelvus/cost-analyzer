import { Workbook } from "@oai/artifact-tool";

const workbook = Workbook.create();
console.log(workbook.help("auto filter", {
  include: "index,examples,notes",
  maxChars: 4000,
}).ndjson);

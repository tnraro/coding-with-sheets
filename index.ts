import * as xlsx from "xlsx";
import { join } from "node:path";
import { $ } from "bun";

const [_, __, path] = Bun.argv;

const buf = await Bun.file(path).arrayBuffer();
const book = xlsx.read(buf);

interface Pos {
  col: number;
  row: number;
}

const cellToPos = (cell: string): Pos[] => {
  return [...cell
    ?.matchAll(/(?<col>[A-Z])(?<row>\d+)/g)]
    .map(x => ({
      col: x.groups!.col.codePointAt(0)! - 65,
      row: Number(x.groups!.row) - 1,
    }))
}
const posToCell = (pos: Pos) => {
  return `${String.fromCodePoint(65 + pos.col)}${pos.row + 1}`;
}

await $`mkdir -p out`;

for (const name of book.SheetNames) {
  const path = name.replaceAll(/_/g, "/");
  const sheet = book.Sheets[name];

  const [from, to] = cellToPos(sheet["!ref"] ?? "");

  const lines: string[] = [];
  for (let y = from.row; y <= to.row; y++) {
    const line: string[] = [];
    for (let x = from.col; x <= to.col; x++) {
      const cell = posToCell({ col: x, row: y });
      const value = sheet[cell]?.v;
      if (value)
        line.push(value);
    }
    lines.push(line.join(" "));
  }
  Bun.write(join("out", path), lines.join("\n"));
}
await $`cd out && mv .run .run.sh && bun .run.sh`;
await $`rm -r out`;

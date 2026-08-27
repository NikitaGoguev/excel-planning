export function buildListsSheet(context) {
  const { addTable, applyHeader, applyInput, applyTitle, setWidths, testData } = context;
  const { lists } = context;
  const artifacts = testData.taskCommentOptions?.artifacts ?? [];
  const adjacentTeams = testData.taskCommentOptions?.adjacentTeams ?? [];

  lists.getRange("A1:C1").merge();
  lists.getRange("A1").values = [["Списки"]];
  applyTitle(lists.getRange("A1:C1"));

  lists.getRange("A3:A7").values = [
    ["Артефакты"],
    ...Array.from({ length: 4 }, (_, index) => [artifacts[index] ?? ""]),
  ];
  lists.getRange("C3:C12").values = [
    ["Смежники"],
    ...Array.from({ length: 9 }, (_, index) => [adjacentTeams[index] ?? ""]),
  ];
  applyHeader(lists.getRange("A3"));
  applyHeader(lists.getRange("C3"));
  applyInput(lists.getRange("A4:A7"));
  applyInput(lists.getRange("C4:C12"));
  addTable(lists, "A3:A7", "tblTaskCommentArtifacts");
  addTable(lists, "C3:C12", "tblTaskCommentAdjacentTeams");

  lists.freezePanes.freezeRows(3);
  setWidths(lists, { A: 260, B: 35, C: 260 });
}

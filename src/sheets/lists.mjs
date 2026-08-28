export function buildListsSheet(context) {
  const { addTable, applyHeader, applyInput, applyTitle, quarterPlanStatuses, setWidths, testData } = context;
  const { lists } = context;
  const artifacts = testData.taskCommentOptions?.artifacts ?? [];
  const adjacentTeams = testData.taskCommentOptions?.adjacentTeams ?? [];

  lists.getRange("A1:G1").merge();
  lists.getRange("A1").values = [["Списки"]];
  applyTitle(lists.getRange("A1:G1"));

  lists.getRange("A3:A7").values = [
    ["Артефакты"],
    ...Array.from({ length: 4 }, (_, index) => [artifacts[index] ?? ""]),
  ];
  lists.getRange("C3:C12").values = [
    ["Смежники"],
    ...Array.from({ length: 9 }, (_, index) => [adjacentTeams[index] ?? ""]),
  ];
  lists.getRange("E3:E10").values = [["Статусы"], ...quarterPlanStatuses.map((status) => [status])];
  lists.getRange("G3:G5").values = [["Да/Нет"], ["Да"], ["Нет"]];
  applyHeader(lists.getRange("A3"));
  applyHeader(lists.getRange("C3"));
  applyHeader(lists.getRange("E3"));
  applyHeader(lists.getRange("G3"));
  applyInput(lists.getRange("A4:A7"));
  applyInput(lists.getRange("C4:C12"));
  applyInput(lists.getRange("E4:E10"));
  applyInput(lists.getRange("G4:G5"));
  addTable(lists, "A3:A7", "tblTaskCommentArtifacts");
  addTable(lists, "C3:C12", "tblTaskCommentAdjacentTeams");
  addTable(lists, "E3:E10", "tblStatuses");
  addTable(lists, "G3:G5", "tblYesNo");

  lists.freezePanes.freezeRows(3);
  setWidths(lists, { A: 260, B: 35, C: 260, D: 35, E: 300, F: 35, G: 110 });
}

import 'package:pluto_grid/pluto_grid.dart';

class ExcelGridData {
  late List<PlutoColumn> columns;
  late List<PlutoRow> rows;

  static const alphabetSize = 26;

  ExcelGridData(List<List> rawRows) {
    columns = rawRows.first
        .asMap()
        .entries
        .map((e) => PlutoColumn(
              title: _getAlphabetTitle(e.key + 1),
              field: e.key.toString(),
              type: PlutoColumnType.text(),
              width: 80,
              readOnly: true,
            ))
        .toList();
    rows = rawRows
        .map((e) => e.asMap().map(
            (key, value) => MapEntry(key.toString(), PlutoCell(value: value))))
        .map((e) => PlutoRow(cells: e))
        .toList();
  }

  String _getAlphabetTitle(int key) {
    final charA = 'A'.codeUnitAt(0);
    String result = '';
    while (key > 0) {
      int mod = (key % alphabetSize);
      if (mod == 0) mod = alphabetSize;
      result = '${String.fromCharCode(mod + charA - 1)}$result';
      key = (key - mod) ~/ alphabetSize;
    }
    return result;
  }
}

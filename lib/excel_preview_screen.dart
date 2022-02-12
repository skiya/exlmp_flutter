import 'package:flutter/material.dart';
import 'package:pluto_grid/pluto_grid.dart';

import 'excel_grid_data.dart';

class ExcelPreviewPage extends StatelessWidget {
  const ExcelPreviewPage({Key? key, required this.previewGridData})
      : super(key: key);

  final ExcelGridData previewGridData;
  static const routeName = '/preview';

  @override
  Widget build(BuildContext context) {
    return Scaffold(
        appBar: AppBar(
          title: const Text('预览'),
          leading: BackButton(onPressed: () {
            Navigator.pop(context);
          }),
        ),
        body: Container(
          padding: const EdgeInsets.all(15),
          child: PlutoGrid(
            columns: previewGridData.columns,
            rows: previewGridData.rows,
          ),
        ));
  }
}

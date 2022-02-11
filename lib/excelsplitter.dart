import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

class ExcelSplitter extends StatefulWidget {
  const ExcelSplitter({Key? key}) : super(key: key);

  @override
  _ExcelSplitterState createState() => _ExcelSplitterState();
}

class _ExcelSplitterState extends State<ExcelSplitter> {
  final GlobalKey<ScaffoldState> _scaffoldKey = GlobalKey<ScaffoldState>();
  final _scaffoldMessengerKey = GlobalKey<ScaffoldMessengerState>();
  String? _fileName;
  String? _saveAsFileName;
  List<PlatformFile>? _paths;
  String? _directoryPath;
  final String? _extension = 'xls, xlsx';
  bool _isLoading = false;
  bool _userAborted = false;
  bool _multiPick = false;
  final FileType _pickingType = FileType.custom;

  void _pickFiles() async {
    _resetState();
    try {
      _directoryPath = null;
      _paths = (await FilePicker.platform.pickFiles(
              type: _pickingType,
              allowMultiple: _multiPick,
              onFileLoading: (FilePickerStatus status) => print(status),
              allowedExtensions: (_extension?.isNotEmpty ?? false)
                  ? _extension?.replaceAll(' ', '').split(',')
                  : null,
              withData: true))
          ?.files;
    } on PlatformException catch (e) {
      _logException('Unsupported operation' + e.toString());
    } catch (e) {
      _logException(e.toString());
    }
    if (!mounted) return;
    setState(() {
      _isLoading = false;
      _fileName =
          _paths != null ? _paths!.map((e) => e.name).toString() : '...';
      _userAborted = _paths == null;
    });
  }

  Future<void> _splitFile() async {
    setState(() {
      _isLoading = true;
    });
    _paths?.forEach((path) {
      final excel = Excel.decodeBytes(path.bytes);
      var resultExcel1 = Excel.createExcel();
      var resultExcel2 = Excel.createExcel();

      for (var table in excel.tables.keys) {
        final allRows = excel.tables[table]?.rows;
        allRows
            ?.where((currRow) => (allRows.indexOf(currRow) + 1).isEven)
            .forEach((element) {
          resultExcel1.appendRow(table, element);
        });
        allRows
            ?.where((currRow) => (allRows.indexOf(currRow)).isEven)
            .forEach((element) {
          resultExcel2.appendRow(table, element);
        });
      }

      resultExcel1.encode().then((value) {
        File('${path.path}_evenRows.${path.extension}')
          ..createSync(recursive: true)
          ..writeAsBytesSync(value);
      });

      resultExcel2.encode().then((value) {
        File('${path.path}_oddRows.${path.extension}')
          ..createSync(recursive: true)
          ..writeAsBytesSync(value);
      });
    });
    setState(() => _isLoading = false);
  }

  void _logException(String message) {
    print(message);
    _scaffoldMessengerKey.currentState?.hideCurrentSnackBar();
    _scaffoldMessengerKey.currentState?.showSnackBar(
      SnackBar(
        content: Text(message),
      ),
    );
  }

  void _resetState() {
    if (!mounted) {
      return;
    }
    setState(() {
      _isLoading = true;
      _directoryPath = null;
      _fileName = null;
      _paths = null;
      _saveAsFileName = null;
      _userAborted = false;
    });
  }

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      scaffoldMessengerKey: _scaffoldMessengerKey,
      home: Scaffold(
        key: _scaffoldKey,
        appBar: AppBar(
          title: const Text('File Picker example app'),
        ),
        body: Center(
          child: Padding(
            padding: const EdgeInsets.only(left: 10.0, right: 10.0),
            child: SingleChildScrollView(
              child: Column(
                mainAxisAlignment: MainAxisAlignment.center,
                children: <Widget>[
                  ConstrainedBox(
                    constraints: const BoxConstraints.tightFor(width: 200.0),
                    child: SwitchListTile.adaptive(
                      title: const Text(
                        'Pick multiple files',
                        textAlign: TextAlign.right,
                      ),
                      onChanged: (bool value) =>
                          setState(() => _multiPick = value),
                      value: _multiPick,
                    ),
                  ),
                  Padding(
                    padding: const EdgeInsets.only(top: 50.0, bottom: 20.0),
                    child: Column(
                      children: <Widget>[
                        ElevatedButton(
                          onPressed: () => _pickFiles(),
                          child: Text(_multiPick ? 'Pick files' : 'Pick file'),
                        ),
                        const SizedBox(height: 10),
                        ElevatedButton(
                          onPressed: () => _splitFile(),
                          child: const Text('Split file'),
                        ),
                        const SizedBox(height: 10),
                      ],
                    ),
                  ),
                  Builder(
                    builder: (BuildContext context) => _isLoading
                        ? const Padding(
                            padding: EdgeInsets.only(bottom: 10.0),
                            child: CircularProgressIndicator(),
                          )
                        : _userAborted
                            ? const Padding(
                                padding: EdgeInsets.only(bottom: 10.0),
                                child: Text(
                                  'User has aborted the dialog',
                                ),
                              )
                            : _directoryPath != null
                                ? ListTile(
                                    title: const Text('Directory path'),
                                    subtitle: Text(_directoryPath!),
                                  )
                                : _paths != null
                                    ? Container(
                                        padding:
                                            const EdgeInsets.only(bottom: 30.0),
                                        height:
                                            MediaQuery.of(context).size.height *
                                                0.50,
                                        child: Scrollbar(
                                            child: ListView.separated(
                                          itemCount: _paths != null &&
                                                  _paths!.isNotEmpty
                                              ? _paths!.length
                                              : 1,
                                          itemBuilder: (BuildContext context,
                                              int index) {
                                            final bool isMultiPath =
                                                _paths != null &&
                                                    _paths!.isNotEmpty;
                                            final String name =
                                                'File $index: ' +
                                                    (isMultiPath
                                                        ? _paths!
                                                            .map((e) => e.name)
                                                            .toList()[index]
                                                        : _fileName ?? '...');
                                            final path = kIsWeb
                                                ? null
                                                : _paths!
                                                    .map((e) => e.path)
                                                    .toList()[index]
                                                    .toString();

                                            return ListTile(
                                              title: Text(
                                                name,
                                              ),
                                              subtitle: Text(path ?? ''),
                                            );
                                          },
                                          separatorBuilder:
                                              (BuildContext context,
                                                      int index) =>
                                                  const Divider(),
                                        )),
                                      )
                                    : _saveAsFileName != null
                                        ? ListTile(
                                            title: const Text('Save file'),
                                            subtitle: Text(_saveAsFileName!),
                                          )
                                        : const SizedBox(),
                  ),
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}

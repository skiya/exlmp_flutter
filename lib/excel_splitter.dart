import 'dart:io';
import 'dart:math';

import 'package:excel/excel.dart';
import 'package:exlmp_flutter/excel_grid_data.dart';
import 'package:exlmp_flutter/excel_preview_screen.dart';
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
  final String? _extension = 'xls, xlsx';
  final FileType _pickingType = FileType.custom;

  String? _fileName;
  List<PlatformFile>? _paths;
  String? _directoryPath;
  var _selectedFileIndice = <int>[];
  bool _isLoading = false;
  bool _userAborted = false;
  bool _multiPick = false;
  bool _splitting = false;
  int _splitInto = 0;
  final _splitPattern = <int, String>{};

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

  // Future<void> _splitFile() async {
  //   setState(() {
  //     _isLoading = true;
  //     _splitting = true;
  //   });
  //   _paths?.forEach((path) {
  //     final excel = Excel.decodeBytes(path.bytes);
  //     var resultExcel1 = Excel.createExcel();
  //     var resultExcel2 = Excel.createExcel();
  //
  //     for (var table in excel.tables.keys) {
  //       final allRows = excel.tables[table]?.rows;
  //       allRows
  //           ?.where((currRow) => (allRows.indexOf(currRow) + 1).isEven)
  //           .forEach((element) {
  //         resultExcel1.appendRow(table, element);
  //       });
  //       allRows
  //           ?.where((currRow) => (allRows.indexOf(currRow)).isEven)
  //           .forEach((element) {
  //         resultExcel2.appendRow(table, element);
  //       });
  //     }
//
  //     resultExcel1.encode().then((value) {
  //       File('${path.path}_evenRows.${path.extension}')
  //         ..createSync(recursive: true)
  //         ..writeAsBytesSync(value);
  //     });
//
  //     resultExcel2.encode().then((value) {
  //       File('${path.path}_oddRows.${path.extension}')
  //         ..createSync(recursive: true)
  //         ..writeAsBytesSync(value);
  //     });
  //   });
  //   setState(() {
  //     _isLoading = false;
  //     _splitting = false;
  //   });
  // }

  Future<void> _updateSelectedList(int index, bool? value) async {
    if (value == true && !_selectedFileIndice.contains(index)) {
      setState(() {
        _selectedFileIndice.add(index);
      });
    } else {
      setState(() {
        _selectedFileIndice.remove(index);
      });
    }
  }

  Future<void> _splitAllFiles() async {
    setState(() {
      _splitting = true;
    });
    final path = await FilePicker.platform.getDirectoryPath();
    if (path == null) {
      setState(() {
        _userAborted = true;
        return;
      });
    }
    final excelsToSplit = _paths!
        .asMap()
        .map((index, file) => MapEntry(file, Excel.decodeBytes(file.bytes)));

    _split(excelsToSplit, path!);
    setState(() {
      _splitting = false;
    });
    _logCompletion("文件已存储在$path");
  }

  Future<void> _splitSelectedFiles() async {
    setState(() {
      _splitting = true;
    });
    final path = await FilePicker.platform.getDirectoryPath();
    if (path == null) {
      setState(() {
        _userAborted = true;
      });
      return;
    }
    final excelsToSplit = Map.fromEntries(_selectedFileIndice
        .map((e) => MapEntry(_paths![e], Excel.decodeBytes(_paths![e].bytes))));
    _split(excelsToSplit, path);
    setState(() {
      _splitting = false;
    });
    _logCompletion("文件已存储在$path");
  }

  void _split(Map<PlatformFile, Excel> excelsToSplit, String path) {
    for (var entry in excelsToSplit.entries) {
      var results = _splitPattern.map(
          (resultIndex, pattern) => MapEntry(pattern, Excel.createExcel()));

      for (var table in entry.value.tables.keys) {
        final sheet = entry.value.tables[table];
        results.forEach((pattern, resultFile) async {
          final List<int> intKeys = getIntKeys(pattern);
          List<List>? rowsToAppend;
          try {
            rowsToAppend = sheet?.rows
                .map((row) => intKeys.map((intKey) => row[intKey]).toList())
                .toList();
          } catch (e) {
            _logException("要选取的列不存在!");
          }

          rowsToAppend?.forEach((row) {
            resultFile.appendRow(table, row);
          });
        });
      }
      results.forEach((pattern, resultFile) {
        resultFile.encode().then((bytes) {
          File('$path/exlmp_${pattern.split(',').join('_')}_${entry.key.name}')
            ..createSync(recursive: true)
            ..writeAsBytesSync(bytes);
        });
      });
    }
  }

  Future<void> _handleSplitIntoInput(String value) async {
    final number = num.tryParse(value);
    if (number != null) {
      setState(() {
        _splitInto = number.toInt();
      });
    }
  }

  Future<void> _handlePatternInput(int index, String value) async {
    setState(() {
      _splitPattern.update(index, (curr) => value, ifAbsent: () => value);
    });
  }

  void _logException(String message) {
    if (kDebugMode) {
      print(message);
    }
    _scaffoldMessengerKey.currentState?.hideCurrentSnackBar();
    _scaffoldMessengerKey.currentState?.showSnackBar(
      SnackBar(
        content: Text(message),
      ),
    );
  }

  void _logCompletion(String message) {
    if (kDebugMode) {
      print(message);
    }
    _scaffoldMessengerKey.currentState?.hideCurrentSnackBar();
    _scaffoldMessengerKey.currentState?.showSnackBar(
      SnackBar(
        content: Text(message),
        backgroundColor: const Color.fromARGB(255, 55, 141, 58),
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
      _userAborted = false;
      _selectedFileIndice = <int>[];
    });
  }

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      scaffoldMessengerKey: _scaffoldMessengerKey,
      onGenerateRoute: (settings) {
        if (settings.name == ExcelPreviewPage.routeName) {
          final args = settings.arguments as ExcelGridData;

          return MaterialPageRoute(
            builder: (context) => ExcelPreviewPage(previewGridData: args),
          );
        }
      },
      home: Scaffold(
        key: _scaffoldKey,
        appBar: AppBar(
          title: const Text('Excel按列分割器'),
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
                    child: Column(
                      children: [
                        TextFormField(
                          decoration: const InputDecoration(
                            labelText: '分割数',
                          ),
                          inputFormatters: [
                            FilteringTextInputFormatter.digitsOnly
                          ],
                          onChanged: (value) {
                            _handleSplitIntoInput(value);
                          },
                        ),
                        ListView.builder(
                            shrinkWrap: true,
                            itemCount: _splitInto,
                            itemBuilder: (context, index) => TextFormField(
                                  decoration: InputDecoration(
                                      labelText: '子文件${index + 1}含有的列',
                                      hintText: '(字母序号以英文逗号隔开)'),
                                  inputFormatters: [
                                    FilteringTextInputFormatter.allow(
                                        RegExp(r'[A-Z,]'))
                                  ],
                                  onChanged: (value) {
                                    _handlePatternInput(index, value);
                                  },
                                )),
                        Padding(
                          padding: const EdgeInsets.only(top: 8.0),
                          child: SwitchListTile.adaptive(
                            title: const Text(
                              '开启此选项以选择多个文件',
                              textAlign: TextAlign.right,
                            ),
                            onChanged: (bool value) =>
                                setState(() => _multiPick = value),
                            value: _multiPick,
                          ),
                        ),
                      ],
                    ),
                  ),
                  Padding(
                    padding: const EdgeInsets.only(top: 8.0, bottom: 20.0),
                    child: Column(
                      children: <Widget>[
                        ElevatedButton(
                          onPressed: () => _pickFiles(),
                          child: Text(_multiPick ? '浏览并选择多个文件' : '浏览并选择文件'),
                        ),
                        const SizedBox(height: 10),
                        ElevatedButton(
                          onPressed: _splitAllFiles,
                          child: const Text('开始分割所有文件'),
                        ),
                        const SizedBox(height: 10),
                        _selectedFileIndice.isNotEmpty
                            ? ElevatedButton(
                                onPressed: _splitSelectedFiles,
                                child: const Text('开始分割选中文件'))
                            : Container(),
                        const SizedBox(height: 10),
                      ],
                    ),
                  ),
                  Builder(
                    builder: (BuildContext context) => _isLoading || _splitting
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
                                            final previewData = ExcelGridData(
                                                Excel.decodeBytes(
                                                        _paths![index].bytes)
                                                    .tables
                                                    .values
                                                    .first
                                                    .rows);

                                            return ListTile(
                                              title: Text(
                                                name,
                                              ),
                                              subtitle: Text(path ?? ''),
                                              leading: Checkbox(
                                                value: _selectedFileIndice
                                                    .contains(index),
                                                onChanged: (value) {
                                                  _updateSelectedList(
                                                      index, value);
                                                },
                                              ),
                                              trailing: IconButton(
                                                  icon:
                                                      const Icon(Icons.preview),
                                                  onPressed: () {
                                                    Navigator.pushNamed(
                                                        context,
                                                        ExcelPreviewPage
                                                            .routeName,
                                                        arguments: previewData);
                                                  }),
                                            );
                                          },
                                          separatorBuilder:
                                              (BuildContext context,
                                                      int index) =>
                                                  const Divider(),
                                        )),
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

List<int> getIntKeys(String keys) {
  final charA = 'A'.codeUnitAt(0);
  final fields = keys.split(',');
  var intKeys = <int>[];
  for (var field in fields) {
    final pows =
        field.codeUnits.map((e) => e - charA + 1).toList().reversed.toList();
    var key = 0;
    for (var i = 0; i < pows.length; i++) {
      key += (pow(26, i) * pows[i]).toInt();
    }
    intKeys.add(key - 1);
  }
  return intKeys;
}

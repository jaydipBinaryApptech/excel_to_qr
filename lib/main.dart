import 'dart:io' show File;
import 'dart:typed_data';
import 'dart:convert' show base64Encode;
import 'package:firebase_core/firebase_core.dart';
import 'package:flutter/cupertino.dart';
import 'package:flutter/foundation.dart' show kIsWeb;
import 'package:flutter/material.dart';
import 'package:flutter/rendering.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'package:barcode_widget/barcode_widget.dart';
import 'dart:ui' as ui;
import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:printing/printing.dart';
import 'package:image_gallery_saver/image_gallery_saver.dart';
// ignore: avoid_web_libraries_in_flutter
import 'dart:html' as html show AnchorElement;

import 'firebase_options.dart';

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized(); // ensures bindings are ready
  await Firebase.initializeApp(options: DefaultFirebaseOptions.currentPlatform);
  runApp(
    const MaterialApp(
      home: ExcelBarcodeGenerator(),
      debugShowCheckedModeBanner: false,
    ),
  );
}

class ExcelBarcodeGenerator extends StatefulWidget {
  const ExcelBarcodeGenerator({Key? key}) : super(key: key);

  @override
  State<ExcelBarcodeGenerator> createState() => _ExcelBarcodeGeneratorState();
}

class _ExcelBarcodeGeneratorState extends State<ExcelBarcodeGenerator> {
  List<String> _cellValues = [];
  List<String> _filteredValues = [];
  final TextEditingController _searchController = TextEditingController();
  final Map<String, GlobalKey> _barcodeKeys = {};

  @override
  void dispose() {
    _searchController.dispose();
    super.dispose();
  }

  Future<void> _pickExcelFile() async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'xls'],
      withData: true,
    );

    if (result == null) return;

    Uint8List? fileBytes;

    if (kIsWeb) {
      fileBytes = result.files.single.bytes;
    } else {
      final path = result.files.single.path;
      if (path != null) {
        fileBytes = File(path).readAsBytesSync();
      }
    }

    if (fileBytes == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Failed to load file bytes')),
      );
      return;
    }

    final excel = Excel.decodeBytes(fileBytes);
    List<String> values = [];

    for (var table in excel.tables.keys) {
      final sheet = excel.tables[table]!;
      for (var row in sheet.rows) {
        for (var cell in row) {
          if (cell?.value != null && cell!.value.toString().trim().isNotEmpty) {
            values.add(cell.value.toString());
          }
        }
      }
    }

    setState(() {
      _cellValues = values;
      _filteredValues = values;
      _searchController.clear();
      // Create keys for each barcode
      _barcodeKeys.clear();
      for (var value in values) {
        _barcodeKeys[value] = GlobalKey();
      }
    });
  }

  void _filterBarcodes(String query) {
    setState(() {
      if (query.isEmpty) {
        _filteredValues = _cellValues;
      } else {
        _filteredValues =
            _cellValues
                .where(
                  (value) => value.toLowerCase().contains(query.toLowerCase()),
                )
                .toList();
      }
    });
  }

  int _getCrossAxisCount(BuildContext context) {
    final width = MediaQuery.of(context).size.width;
    if (width > 1200) return 4;
    if (width > 800) return 3;
    if (width > 600) return 2;
    return 1;
  }

  Future<void> _downloadAllAsPDF() async {
    if (_cellValues.isEmpty) {
      ScaffoldMessenger.of(
        context,
      ).showSnackBar(const SnackBar(content: Text('No barcodes to download')));
      return;
    }

    final pdf = pw.Document();

    // Add barcodes to PDF in grid format
    for (int i = 0; i < _cellValues.length; i += 12) {
      final chunk = _cellValues.skip(i).take(12).toList();

      pdf.addPage(
        pw.Page(
          pageFormat: PdfPageFormat.a4,
          build: (context) {
            return pw.Wrap(
              spacing: 8,
              runSpacing: 8,
              children:
                  chunk.map((value) {
                    return pw.Container(
                      width: 180,
                      height: 60,
                      decoration: pw.BoxDecoration(
                        border: pw.Border.all(color: PdfColors.grey300),
                      ),
                      child: pw.Column(
                        mainAxisAlignment: pw.MainAxisAlignment.center,
                        children: [
                          pw.Text(
                            value,
                            style: pw.TextStyle(
                              fontSize: 12,
                              fontWeight: pw.FontWeight.bold,
                            ),
                          ),
                          pw.SizedBox(height: 4),
                          pw.BarcodeWidget(
                            barcode: pw.Barcode.code128(),
                            data: value,
                            width: 160,
                            height: 35,
                            drawText: false,
                          ),
                        ],
                      ),
                    );
                  }).toList(),
            );
          },
        ),
      );
    }

    // Save/Print PDF
    await Printing.layoutPdf(onLayout: (format) async => pdf.save());
  }

  Future<void> _downloadSingleBarcodeAsImage(String value) async {
    try {
      final key = _barcodeKeys[value];
      if (key == null || key.currentContext == null) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Failed to capture barcode')),
        );
        return;
      }

      // Find the RenderRepaintBoundary
      RenderRepaintBoundary boundary =
          key.currentContext!.findRenderObject() as RenderRepaintBoundary;

      // Capture the image with high quality
      ui.Image image = await boundary.toImage(pixelRatio: 3.0);
      ByteData? byteData = await image.toByteData(
        format: ui.ImageByteFormat.png,
      );

      if (byteData == null) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Failed to generate image')),
        );
        return;
      }

      Uint8List pngBytes = byteData.buffer.asUint8List();

      if (kIsWeb) {
        // For web, trigger download
        final anchor =
            html.AnchorElement(
                href: 'data:image/png;base64,${base64Encode(pngBytes)}',
              )
              ..setAttribute('download', 'barcode_$value.png')
              ..click();
      } else {
        // For mobile/desktop, save to gallery
        final result = await ImageGallerySaver.saveImage(
          pngBytes,
          name: 'barcode_$value',
          quality: 100,
        );

        if (result['isSuccess']) {
          ScaffoldMessenger.of(context).showSnackBar(
            const SnackBar(content: Text('Barcode saved to gallery')),
          );
        } else {
          ScaffoldMessenger.of(context).showSnackBar(
            const SnackBar(content: Text('Failed to save barcode')),
          );
        }
      }
    } catch (e) {
      ScaffoldMessenger.of(
        context,
      ).showSnackBar(SnackBar(content: Text('Error: $e')));
    }
  }

  @override
  Widget build(BuildContext context) {
    final crossAxisCount = _getCrossAxisCount(context);

    return Scaffold(
      backgroundColor: CupertinoColors.white,
      appBar: AppBar(
        backgroundColor: Colors.blueAccent, // Bright color for contrast
        titleSpacing: 0,
        title: Row(
          children: [
            // const SizedBox(width: 12),
            // if (_cellValues.isEmpty && )
            //   const Text(
            //     'Excel → Barcode Generator',
            //     style: TextStyle(
            //       color: Colors.white,
            //       fontSize: 15,
            //       fontWeight: FontWeight.bold,
            //     ),
            //   ),
            // const Spacer(),

            // 🔍 Search Bar (visible after upload)
            if (_cellValues.isNotEmpty)
              Expanded(
                flex: 3,
                child: Padding(
                  padding: const EdgeInsets.symmetric(
                    vertical: 8,
                    horizontal: 8,
                  ),
                  child: TextField(
                    controller: _searchController,
                    onChanged: _filterBarcodes,
                    style: const TextStyle(color: Colors.black, fontSize: 14),
                    decoration: InputDecoration(
                      hintText: 'Search barcodes...',
                      hintStyle: const TextStyle(color: Colors.black54),
                      prefixIcon: const Icon(
                        Icons.search,
                        color: Colors.black54,
                      ),
                      suffixIcon:
                          _searchController.text.isNotEmpty
                              ? IconButton(
                                icon: const Icon(
                                  Icons.clear,
                                  color: Colors.black54,
                                ),
                                onPressed: () {
                                  _searchController.clear();
                                  _filterBarcodes('');
                                },
                              )
                              : null,
                      filled: true,
                      fillColor: Colors.white, // clear white for visibility
                      border: OutlineInputBorder(
                        borderRadius: BorderRadius.circular(30),
                        borderSide: BorderSide.none,
                      ),
                      contentPadding: const EdgeInsets.symmetric(
                        horizontal: 16,
                        vertical: 0,
                      ),
                    ),
                  ),
                ),
              ),

            const SizedBox(width: 8),

            // 📁 Upload Excel Button (with text for clarity)
            ElevatedButton.icon(
              style: ElevatedButton.styleFrom(
                backgroundColor: Colors.white,
                foregroundColor: Colors.blueAccent,
                padding: const EdgeInsets.symmetric(
                  horizontal: 12,
                  vertical: 10,
                ),
                shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(30),
                ),
                elevation: 0,
              ),
              onPressed: _pickExcelFile,
              icon: const Icon(Icons.upload_file, size: 20),
              label: const Text(
                'Upload Excel',
                style: TextStyle(fontWeight: FontWeight.bold),
              ),
            ),

            // const SizedBox(width: 12),
            //
            // // ⬇️ Download All button
            // if (_cellValues.isNotEmpty)
            //   ElevatedButton.icon(
            //     style: ElevatedButton.styleFrom(
            //       backgroundColor: Colors.white,
            //       foregroundColor: Colors.green,
            //       padding: const EdgeInsets.symmetric(
            //         horizontal: 12,
            //         vertical: 10,
            //       ),
            //       shape: RoundedRectangleBorder(
            //         borderRadius: BorderRadius.circular(30),
            //       ),
            //       elevation: 0,
            //     ),
            //     onPressed: _downloadAllAsPDF,
            //     icon: const Icon(Icons.download, size: 20),
            //     label: const Text(
            //       'Download All',
            //       style: TextStyle(fontWeight: FontWeight.bold),
            //     ),
            //   ),
            const SizedBox(width: 12),
          ],
        ),
      ),

      body: Padding(
        padding: const EdgeInsets.all(12.0),
        child: Column(
          children: [
            // ElevatedButton.icon(
            //   onPressed: _pickExcelFile,
            //   icon: const Icon(Icons.upload_file),
            //   label: const Text('Upload Excel File'),
            // ),
            // if (_cellValues.isNotEmpty) ...[
            //   const SizedBox(height: 12),
            //   TextField(
            //     controller: _searchController,
            //     decoration: InputDecoration(
            //       hintText: 'Search barcodes...',
            //       prefixIcon: const Icon(Icons.search),
            //       suffixIcon:
            //           _searchController.text.isNotEmpty
            //               ? IconButton(
            //                 icon: const Icon(Icons.clear),
            //                 onPressed: () {
            //                   _searchController.clear();
            //                   _filterBarcodes('');
            //                 },
            //               )
            //               : null,
            //       border: OutlineInputBorder(
            //         borderRadius: BorderRadius.circular(8),
            //       ),
            //       contentPadding: const EdgeInsets.symmetric(
            //         horizontal: 16,
            //         vertical: 12,
            //       ),
            //     ),
            //     onChanged: _filterBarcodes,
            //   ),
            // ],
            const SizedBox(height: 12),
            Expanded(
              child:
                  _cellValues.isEmpty
                      ? const Center(
                        child: Text('Upload Excel to see barcodes'),
                      )
                      : _filteredValues.isEmpty
                      ? const Center(
                        child: Text('No barcodes match your search'),
                      )
                      : GridView.builder(
                        gridDelegate: SliverGridDelegateWithFixedCrossAxisCount(
                          crossAxisCount: crossAxisCount,
                          childAspectRatio: 3 / 1.2,
                          crossAxisSpacing: 12,
                          mainAxisSpacing: 12,
                        ),
                        itemCount: _filteredValues.length,
                        itemBuilder: (context, index) {
                          final value = _filteredValues[index];
                          return Card(
                            elevation: 3,
                            child: Stack(
                              children: [
                                RepaintBoundary(
                                  key: _barcodeKeys[value],
                                  child: Container(
                                    color: Colors.white,
                                    child: Padding(
                                      padding: const EdgeInsets.symmetric(
                                        horizontal: 8,
                                        vertical: 6,
                                      ),
                                      child: LayoutBuilder(
                                        builder: (context, constraints) {
                                          double textSize =
                                              constraints.maxWidth *
                                              0.12; // scales with card width
                                          textSize = textSize.clamp(
                                            14,
                                            28,
                                          ); // min/max font size limit

                                          return Column(
                                            mainAxisAlignment:
                                                MainAxisAlignment.center,
                                            children: [
                                              // Barcode text (same style as image)
                                              Padding(
                                                padding: const EdgeInsets.only(
                                                  bottom: 0,
                                                ),
                                                child: Text(
                                                  value,
                                                  textAlign: TextAlign.center,
                                                  style: TextStyle(
                                                    fontSize: textSize,
                                                    fontWeight: FontWeight.w700,
                                                    letterSpacing: 1.2,
                                                  ),
                                                ),
                                              ),

                                              // Barcode image with fixed height ratio
                                              Expanded(
                                                child: Center(
                                                  child: BarcodeWidget(
                                                    barcode: Barcode.code128(),
                                                    data: value,
                                                    drawText: false,
                                                    width:
                                                        constraints.maxWidth *
                                                        0.9,
                                                    height:
                                                        constraints.maxHeight *
                                                        0.55,
                                                  ),
                                                ),
                                              ),
                                            ],
                                          );
                                        },
                                      ),
                                    ),
                                  ),
                                ),
                                Positioned(
                                  top: 4,
                                  right: 4,
                                  child: IconButton(
                                    icon: const Icon(Icons.download, size: 24),
                                    padding: EdgeInsets.zero,
                                    constraints: const BoxConstraints(),
                                    onPressed:
                                        () => _downloadSingleBarcodeAsImage(
                                          value,
                                        ),
                                    tooltip: 'Download as image',
                                  ),
                                ),
                              ],
                            ),
                          );
                        },
                      ),
            ),
          ],
        ),
      ),
    );
  }
}

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
import 'package:google_fonts/google_fonts.dart';
import 'dart:ui' as ui;
import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:printing/printing.dart';
import 'package:image_gallery_saver/image_gallery_saver.dart';
import 'package:archive/archive.dart';
// ignore: avoid_web_libraries_in_flutter
import 'dart:html' as html show AnchorElement;

import 'firebase_options.dart';

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
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
  bool _isDownloading = false;

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

  Future<Uint8List?> _captureBarcodeAsImage(String value) async {
    try {
      final key = _barcodeKeys[value];
      if (key == null || key.currentContext == null) {
        return null;
      }

      RenderRepaintBoundary boundary =
          key.currentContext!.findRenderObject() as RenderRepaintBoundary;

      ui.Image image = await boundary.toImage(pixelRatio: 3.0);
      ByteData? byteData = await image.toByteData(
        format: ui.ImageByteFormat.png,
      );

      if (byteData == null) {
        return null;
      }

      return byteData.buffer.asUint8List();
    } catch (e) {
      print('Error capturing barcode: $e');
      return null;
    }
  }

  Future<void> _downloadAllAsZip() async {
    if (_cellValues.isEmpty) {
      ScaffoldMessenger.of(
        context,
      ).showSnackBar(const SnackBar(content: Text('No barcodes to download')));
      return;
    }

    setState(() {
      _isDownloading = true;
    });

    try {
      // Create a ZIP archive
      final archive = Archive();

      // Capture all barcodes
      for (int i = 0; i < _cellValues.length; i++) {
        final value = _cellValues[i];
        final imageBytes = await _captureBarcodeAsImage(value);

        if (imageBytes != null) {
          // Add image to archive
          final fileName =
              'barcode_${value.replaceAll(RegExp(r'[^\w\s-]'), '_')}.png';
          archive.addFile(ArchiveFile(fileName, imageBytes.length, imageBytes));
        }

        // Show progress
        if ((i + 1) % 10 == 0 || i == _cellValues.length - 1) {
          ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(
              content: Text('Processing: ${i + 1}/${_cellValues.length}'),
              duration: const Duration(milliseconds: 500),
            ),
          );
        }
      }

      // Encode the archive as ZIP
      final zipBytes = ZipEncoder().encode(archive);

      if (zipBytes == null) {
        throw Exception('Failed to create ZIP file');
      }

      // Download the ZIP file
      if (kIsWeb) {
        final anchor =
            html.AnchorElement(
                href: 'data:application/zip;base64,${base64Encode(zipBytes)}',
              )
              ..setAttribute(
                'download',
                'barcodes_${DateTime.now().millisecondsSinceEpoch}.zip',
              )
              ..click();

        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('ZIP file downloaded successfully!')),
        );
      } else {
        // For mobile, you might want to use path_provider to save to documents
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text('ZIP download on mobile requires additional setup'),
          ),
        );
      }
    } catch (e) {
      ScaffoldMessenger.of(
        context,
      ).showSnackBar(SnackBar(content: Text('Error creating ZIP: $e')));
    } finally {
      setState(() {
        _isDownloading = false;
      });
    }
  }

  Future<void> _downloadSingleBarcodeAsImage(String value) async {
    try {
      final pngBytes = await _captureBarcodeAsImage(value);

      if (pngBytes == null) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Failed to capture barcode')),
        );
        return;
      }

      if (kIsWeb) {
        final anchor =
            html.AnchorElement(
                href: 'data:image/png;base64,${base64Encode(pngBytes)}',
              )
              ..setAttribute('download', 'barcode_$value.png')
              ..click();
      } else {
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
        backgroundColor: Colors.blueAccent,
        titleSpacing: 0,
        title: Row(
          children: [
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
                      fillColor: Colors.white,
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
            const SizedBox(width: 12),
            if (_cellValues.isNotEmpty)
              ElevatedButton.icon(
                style: ElevatedButton.styleFrom(
                  backgroundColor: Colors.white,
                  foregroundColor: Colors.green,
                  padding: const EdgeInsets.symmetric(
                    horizontal: 12,
                    vertical: 10,
                  ),
                  shape: RoundedRectangleBorder(
                    borderRadius: BorderRadius.circular(30),
                  ),
                  elevation: 0,
                ),
                onPressed: _isDownloading ? null : _downloadAllAsZip,
                icon:
                    _isDownloading
                        ? const SizedBox(
                          width: 20,
                          height: 20,
                          child: CircularProgressIndicator(strokeWidth: 2),
                        )
                        : const Icon(Icons.download, size: 20),
                label: Text(
                  _isDownloading ? 'Processing...' : 'Download All',
                  style: const TextStyle(fontWeight: FontWeight.bold),
                ),
              ),
            const SizedBox(width: 12),
          ],
        ),
      ),
      body: Padding(
        padding: const EdgeInsets.all(12.0),
        child: Column(
          children: [
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
                          childAspectRatio: 14 / 5,
                          crossAxisSpacing: 12,
                          mainAxisSpacing: 12,
                        ),
                        itemCount: _filteredValues.length,
                        itemBuilder: (context, index) {
                          final value = _filteredValues[index];
                          return Card(
                            elevation: 3,
                            clipBehavior: Clip.antiAlias,
                            child: Stack(
                              children: [
                                RepaintBoundary(
                                  key: _barcodeKeys[value],
                                  child: Container(
                                    color: Colors.white,
                                    padding: const EdgeInsets.symmetric(
                                      horizontal: 8,
                                      vertical: 8,
                                    ),
                                    child: LayoutBuilder(
                                      builder: (context, constraints) {
                                        // Calculate available space
                                        final availableHeight =
                                            constraints.maxHeight;
                                        final availableWidth =
                                            constraints.maxWidth;

                                        // Calculate barcode with perfect 14:4 aspect ratio
                                        // Aspect ratio 14:4 means width:height = 14:4 = 3.5:1
                                        // So width = height * 3.5

                                        // Try to fit within available space
                                        double barcodeHeight;
                                        double barcodeWidth;

                                        // Calculate based on width constraint
                                        final maxWidthBasedHeight =
                                            availableWidth / 3.5;
                                        // Calculate based on height constraint
                                        final maxHeightBasedWidth =
                                            availableHeight * 3.5;

                                        if (maxWidthBasedHeight <=
                                            availableHeight) {
                                          // Width is the limiting factor
                                          barcodeHeight = maxWidthBasedHeight;
                                          barcodeWidth = availableWidth;
                                        } else {
                                          // Height is the limiting factor
                                          barcodeHeight = availableHeight;
                                          barcodeWidth = maxHeightBasedWidth;

                                          // If width exceeds available space, scale down
                                          if (barcodeWidth > availableWidth) {
                                            barcodeWidth = availableWidth;
                                            barcodeHeight = barcodeWidth / 3.5;
                                          }
                                        }

                                        // Text size: 20% of current (which was 10% of barcode height)
                                        // So new text size = 2% of barcode height
                                        final textSize = barcodeHeight * 0.02;

                                        return Column(
                                          mainAxisSize: MainAxisSize.min,
                                          mainAxisAlignment:
                                              MainAxisAlignment.center,
                                          children: [
                                            // Barcode with perfect 14:4 aspect ratio
                                            SizedBox(
                                              height: barcodeHeight,
                                              width: barcodeWidth,
                                              child: BarcodeWidget(
                                                barcode: Barcode.code128(),
                                                data: value,
                                                drawText: false,
                                              ),
                                            ),
                                            // No space between barcode and text
                                            // Text below barcode - very small font (20% of previous size)
                                            Text(
                                              value,
                                              textAlign: TextAlign.center,
                                              style: GoogleFonts.robotoMono(
                                                fontSize: 10,
                                                fontWeight: FontWeight.w900,
                                                letterSpacing: 0.001,
                                                color: Colors.black,
                                              ),
                                              maxLines: 1,
                                              overflow: TextOverflow.ellipsis,
                                            ),
                                          ],
                                        );
                                      },
                                    ),
                                  ),
                                ),
                                Positioned(
                                  top: 4,
                                  right: 4,
                                  child: Container(
                                    decoration: BoxDecoration(
                                      color: Colors.green,
                                      shape: BoxShape.circle,
                                      boxShadow: [
                                        BoxShadow(
                                          color: Colors.green.withOpacity(0.5),
                                          blurRadius: 6,
                                          spreadRadius: 1,
                                          offset: const Offset(0, 2),
                                        ),
                                      ],
                                    ),
                                    child: IconButton(
                                      icon: const Icon(
                                        Icons.download,
                                        size: 20,
                                        color: Colors.white,
                                      ),
                                      padding: EdgeInsets.zero,
                                      constraints: const BoxConstraints(
                                        minWidth: 36,
                                        minHeight: 36,
                                      ),
                                      onPressed:
                                          () => _downloadSingleBarcodeAsImage(
                                            value,
                                          ),
                                      tooltip: 'Download as image',
                                    ),
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

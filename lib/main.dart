// PortSafe — safety officer app (test build).
// Sections: Toolbox Talk, Attendance, Container Moves, Damage Report.
// Data is stored on the phone only (sqflite). Fully English UI.
import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';
import 'dart:ui' as ui;

import 'package:excel/excel.dart' hide Border;
import 'package:flutter/material.dart';
import 'package:image_picker/image_picker.dart';
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';
import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:sqflite/sqflite.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const PortSafeApp());
}

String fmtDate(DateTime d) =>
    '${d.year}-${d.month.toString().padLeft(2, '0')}-${d.day.toString().padLeft(2, '0')}';

DateTime parseDate(String s) => DateTime.parse(s);

// ---------------- Day-wise TBT topics (client sheet points) ----------------
const Map<int, Map<String, String>> kDefaultTopics = {
  1: {
    'title': 'Safe Lifting and Manual Handling',
    'detail':
        'Plan every lift before you start. Keep the load close to your body, bend your knees and not your back, and never twist while carrying. Use team lifts for heavy or awkward loads and mechanical aids where available. Report damaged slings, hooks and chains immediately.'
  },
  2: {
    'title': 'Forklift Safety and Blind Spots',
    'detail':
        'Only trained and authorised operators drive forklifts. No person may stay inside a container while a forklift enters it. Stay clear of blind spots, keep a safe distance and make eye contact with the operator before passing. Seat belts must always be worn. Never ride on forks or loads.'
  },
  3: {
    'title': 'Load Distribution, Lashing and Unlashing',
    'detail':
        'Distribute weight evenly and secure cargo before any move. Use the buddy system for lashing and unlashing. Check tension belts for cuts and wear, release tension carefully and stand clear of recoil. Never stand under a suspended load.'
  },
  4: {
    'title': 'Working at Height - Container Top Safety',
    'detail':
        'Use the man cage with a full fall-arrest harness when working on top of containers. Inspect the harness and lanyard before use, keep three points of contact on ladders, and never work on a wet, damaged or unstable container top.'
  },
  5: {
    'title': 'Heat Stress Prevention',
    'detail':
        'Drink water regularly, take shade breaks and watch for dizziness, nausea and confusion in yourself and others. Wear PPE correctly even in heat, report heat illness early, and know the emergency response for heat stroke.'
  },
  6: {
    'title': 'PPE Compliance and Housekeeping',
    'detail':
        'Wear your helmet, safety shoes, gloves and high-visibility vest at all times in the terminal. Keep walkways clear, clean spills immediately, stack materials safely and dispose of waste properly.'
  },
  7: {
    'title': 'Site Inspection and Safe Work Planning',
    'detail':
        'Inspect the work area before starting, check tools and equipment, confirm the job scope and hazards with the foreman, and stop work immediately if conditions are unsafe.'
  },
};

const List<String> kWeekdayNames = [
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday',
  'Sunday'
];

// ---------------- Database ----------------
class DbHelper {
  DbHelper._();
  static final DbHelper instance = DbHelper._();
  Database? _db;

  Future<Database> get db async {
    if (_db != null) return _db!;
    final dir = await getDatabasesPath();
    _db = await openDatabase(
      p.join(dir, 'portsafe.db'),
      version: 1,
      onCreate: (db, version) async {
        await db.execute(
            'CREATE TABLE settings(key TEXT PRIMARY KEY, value TEXT)');
        await db.execute(
            'CREATE TABLE staff(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, designation TEXT, nationality TEXT, shift TEXT, iqama TEXT)');
        await db.execute(
            'CREATE TABLE moves(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, container TEXT, shift TEXT, foreman TEXT)');
        await db.execute(
            'CREATE TABLE move_crew(id INTEGER PRIMARY KEY AUTOINCREMENT, move_id INTEGER, staff_id INTEGER, staff_name TEXT)');
        await db.execute(
            'CREATE TABLE damage(id INTEGER PRIMARY KEY AUTOINCREMENT, container TEXT, date TEXT, shift TEXT, foreman TEXT, labour TEXT, origin TEXT, client_report TEXT, client_report_date TEXT, photo_path TEXT, notes TEXT)');
        await db.execute(
            'CREATE TABLE tbt(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, project TEXT, contractor TEXT, terminal TEXT, shift TEXT, location TEXT, topic TEXT, topic_detail TEXT, photo_path TEXT)');
        await db.execute(
            'CREATE TABLE tbt_attendees(id INTEGER PRIMARY KEY AUTOINCREMENT, tbt_id INTEGER, name TEXT, iqama TEXT, designation TEXT, nationality TEXT, shift TEXT, remarks TEXT, signature BLOB)');
        await db.execute(
            'CREATE TABLE attendance(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, staff_id INTEGER, staff_name TEXT, designation TEXT, shift TEXT, remarks TEXT, ot_hours REAL, signature BLOB)');
      },
    );
    return _db!;
  }

  Future<String> getSetting(String key, [String fallback = '']) async {
    final d = await db;
    final rows = await d.query('settings',
        where: 'key = ?', whereArgs: [key], limit: 1);
    if (rows.isEmpty) return fallback;
    return (rows.first['value'] as String?) ?? fallback;
  }

  Future<void> setSetting(String key, String value) async {
    final d = await db;
    await d.insert('settings', {'key': key, 'value': value},
        conflictAlgorithm: ConflictAlgorithm.replace);
  }

  Future<String> companyName() => getSetting('company_name');
  Future<Uint8List?> logoBytes() async {
    final b64 = await getSetting('company_logo');
    if (b64.isEmpty) return null;
    try {
      return base64Decode(b64);
    } catch (_) {
      return null;
    }
  }

  Future<Map<int, Map<String, String>>> topics() async {
    final raw = await getSetting('topics_json');
    if (raw.isEmpty) {
      return Map<int, Map<String, String>>.from(kDefaultTopics);
    }
    try {
      final decoded = jsonDecode(raw) as Map<String, dynamic>;
      final out = <int, Map<String, String>>{};
      decoded.forEach((k, v) {
        final m = v as Map<String, dynamic>;
        out[int.parse(k)] = {
          'title': (m['title'] as String?) ?? '',
          'detail': (m['detail'] as String?) ?? '',
        };
      });
      for (final e in kDefaultTopics.entries) {
        out.putIfAbsent(e.key, () => e.value);
      }
      return out;
    } catch (_) {
      return Map<int, Map<String, String>>.from(kDefaultTopics);
    }
  }

  Future<List<Map<String, Object?>>> staffList() async {
    final d = await db;
    return d.query('staff', orderBy: 'name ASC');
  }

  Future<int> insertStaff(Map<String, Object?> row) async {
    final d = await db;
    return d.insert('staff', row);
  }

  Future<void> updateStaff(int id, Map<String, Object?> row) async {
    final d = await db;
    await d.update('staff', row, where: 'id = ?', whereArgs: [id]);
  }

  Future<void> deleteStaff(int id) async {
    final d = await db;
    await d.delete('staff', where: 'id = ?', whereArgs: [id]);
  }
}

Future<Directory> appDocs() => getApplicationDocumentsDirectory();

Future<String> saveBytesToDocs(Uint8List bytes, String name) async {
  final dir = await appDocs();
  final file = File(p.join(dir.path, name));
  await file.writeAsBytes(bytes);
  return file.path;
}

Future<String?> pickImagePath() async {
  final f = await ImagePicker().pickImage(source: ImageSource.gallery);
  if (f == null) return null;
  final bytes = await f.readAsBytes();
  final name =
      'img_${DateTime.now().millisecondsSinceEpoch}${p.extension(f.path)}';
  return saveBytesToDocs(bytes, name);
}

// ---------------- Signature pad ----------------
class SignaturePad extends StatefulWidget {
  final void Function(List<Offset?> points) onChanged;
  const SignaturePad({super.key, required this.onChanged});
  @override
  State<SignaturePad> createState() => _SignaturePadState();
}

class _SignaturePadState extends State<SignaturePad> {
  final List<Offset?> _points = [];
  @override
  Widget build(BuildContext context) {
    return Container(
      width: 280,
      height: 130,
      decoration: BoxDecoration(
        border: Border.all(color: Colors.grey),
        color: Colors.white,
      ),
      child: GestureDetector(
        onPanUpdate: (d) {
          final box = context.findRenderObject() as RenderBox;
          final pos = box.globalToLocal(d.globalPosition);
          if (pos.dx >= 0 && pos.dx <= 280 && pos.dy >= 0 && pos.dy <= 130) {
            setState(() => _points.add(pos));
            widget.onChanged(List<Offset?>.from(_points));
          }
        },
        onPanEnd: (_) {
          setState(() => _points.add(null));
          widget.onChanged(List<Offset?>.from(_points));
        },
        child: CustomPaint(
          painter: _SigPainter(_points),
          size: const Size(280, 130),
        ),
      ),
    );
  }
}

class _SigPainter extends CustomPainter {
  final List<Offset?> points;
  _SigPainter(this.points);
  @override
  void paint(Canvas canvas, Size size) {
    final paint = Paint()
      ..color = Colors.black
      ..strokeWidth = 2
      ..strokeCap = StrokeCap.round;
    for (var i = 0; i < points.length - 1; i++) {
      if (points[i] != null && points[i + 1] != null) {
        canvas.drawLine(points[i]!, points[i + 1]!, paint);
      }
    }
  }

  @override
  bool shouldRepaint(_SigPainter old) => old.points != points;
}

Future<Uint8List?> signatureToPng(List<Offset?> points) async {
  if (points.whereType<Offset>().isEmpty) return null;
  final recorder = ui.PictureRecorder();
  final canvas = Canvas(recorder, const Rect.fromLTWH(0, 0, 280, 130));
  canvas.drawRect(
      const Rect.fromLTWH(0, 0, 280, 130), Paint()..color = Colors.white);
  final paint = Paint()
    ..color = Colors.black
    ..strokeWidth = 2
    ..strokeCap = StrokeCap.round;
  for (var i = 0; i < points.length - 1; i++) {
    if (points[i] != null && points[i + 1] != null) {
      canvas.drawLine(points[i]!, points[i + 1]!, paint);
    }
  }
  final picture = recorder.endRecording();
  final img = await picture.toImage(280, 130);
  final data = await img.toByteData(format: ui.ImageByteFormat.png);
  return data?.buffer.asUint8List();
}

Future<Uint8List?> askSignature(BuildContext context) async {
  List<Offset?> points = [];
  final ok = await showDialog<bool>(
    context: context,
    builder: (ctx) => AlertDialog(
      title: const Text('Sign with your finger'),
      content: Column(
        mainAxisSize: MainAxisSize.min,
        children: [
          SignaturePad(onChanged: (pts) => points = pts),
          const SizedBox(height: 8),
          const Text('Draw your signature above, then tap Save.',
              style: TextStyle(fontSize: 12)),
        ],
      ),
      actions: [
        TextButton(
            onPressed: () => Navigator.pop(ctx, false),
            child: const Text('Cancel')),
        ElevatedButton(
            onPressed: () => Navigator.pop(ctx, true),
            child: const Text('Save')),
      ],
    ),
  );
  if (ok != true) return null;
  return signatureToPng(points);
}

void showMsg(BuildContext context, String msg) {
  ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(msg)));
}

// ---------------- App + Home ----------------
class PortSafeApp extends StatelessWidget {
  const PortSafeApp({super.key});
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'PortSafe',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        useMaterial3: true,
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.teal),
      ),
      home: const HomeScreen(),
    );
  }
}

class HomeScreen extends StatefulWidget {
  const HomeScreen({super.key});
  @override
  State<HomeScreen> createState() => _HomeScreenState();
}

class _HomeScreenState extends State<HomeScreen> {
  String _company = '';
  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    final c = await DbHelper.instance.companyName();
    if (mounted) setState(() => _company = c);
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: _company.isEmpty
            ? TextButton(
                onPressed: () async {
                  await Navigator.push(
                      context,
                      MaterialPageRoute(
                          builder: (_) => const SettingsScreen()));
                  _load();
                },
                child: const Text('Add your company name in Settings',
                    style: TextStyle(color: Colors.white)),
              )
            : Text(_company,
                style: const TextStyle(fontWeight: FontWeight.bold)),
        backgroundColor: Colors.teal.shade700,
        foregroundColor: Colors.white,
        actions: [
          IconButton(
            icon: const Icon(Icons.settings),
            tooltip: 'Settings',
            onPressed: () async {
              await Navigator.push(
                  context,
                  MaterialPageRoute(
                      builder: (_) => const SettingsScreen()));
              _load();
            },
          ),
        ],
      ),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          _card(context, Icons.campaign, 'Toolbox Talk (TBT)',
              'Daily talk, client-format sheet, attendees + signatures',
              () async {
            await Navigator.push(context,
                MaterialPageRoute(builder: (_) => const TbtScreen()));
          }),
          _card(context, Icons.checklist, 'Attendance',
              'Daily attendance with signatures and overtime hours',
              () async {
            await Navigator.push(
                context,
                MaterialPageRoute(builder: (_) => const AttendanceScreen()));
          }),
          _card(context, Icons.forklift, 'Container Moves',
              'Stripping and Stuffing moves, crew by tap, daily counts',
              () async {
            await Navigator.push(context,
                MaterialPageRoute(builder: (_) => const MovesScreen()));
          }),
          _card(context, Icons.report_problem, 'Damage Report',
              'Container damage: origin, client report status, photo',
              () async {
            await Navigator.push(context,
                MaterialPageRoute(builder: (_) => const DamageScreen()));
          }),
          _card(context, Icons.people, 'Staff List',
              'Operators, labour and foremen + photo-assisted import',
              () async {
            await Navigator.push(context,
                MaterialPageRoute(builder: (_) => const StaffScreen()));
          }),
          _card(context, Icons.file_download, 'Export TBT',
              'Date range → PDF official record and Excel table', () async {
            await Navigator.push(context,
                MaterialPageRoute(builder: (_) => const ExportScreen()));
          }),
          const SizedBox(height: 16),
          const Text(
            'PortSafe test build. Data is stored on this phone only. No login or online sync in this version.',
            style: TextStyle(fontSize: 12, color: Colors.grey),
            textAlign: TextAlign.center,
          ),
        ],
      ),
    );
  }

  Widget _card(BuildContext context, IconData icon, String title,
      String sub, VoidCallback onTap) {
    return Card(
      margin: const EdgeInsets.only(bottom: 12),
      child: ListTile(
        leading: Icon(icon, color: Colors.teal.shade700, size: 32),
        title: Text(title, style: const TextStyle(fontWeight: FontWeight.bold)),
        subtitle: Text(sub),
        trailing: const Icon(Icons.chevron_right),
        onTap: onTap,
      ),
    );
  }
}

// ---------------- Settings ----------------
class SettingsScreen extends StatefulWidget {
  const SettingsScreen({super.key});
  @override
  State<SettingsScreen> createState() => _SettingsScreenState();
}

class _SettingsScreenState extends State<SettingsScreen> {
  final _companyCtrl = TextEditingController();
  Uint8List? _logo;
  Map<int, Map<String, String>> _topics = {};
  bool _loaded = false;

  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    final c = await DbHelper.instance.companyName();
    final logo = await DbHelper.instance.logoBytes();
    final topics = await DbHelper.instance.topics();
    if (mounted) {
      setState(() {
        _companyCtrl.text = c;
        _logo = logo;
        _topics = topics;
        _loaded = true;
      });
    }
  }

  @override
  void dispose() {
    _companyCtrl.dispose();
    super.dispose();
  }

  Future<void> _save() async {
    await DbHelper.instance.setSetting('company_name', _companyCtrl.text);
    final map = <String, dynamic>{};
    _topics.forEach((k, v) {
      map[k.toString()] = v;
    });
    await DbHelper.instance.setSetting('topics_json', jsonEncode(map));
    if (mounted) showMsg(context, 'Settings saved');
    if (mounted) Navigator.pop(context);
  }

  @override
  Widget build(BuildContext context) {
    if (!_loaded) {
      return const Scaffold(body: Center(child: CircularProgressIndicator()));
    }
    return Scaffold(
      appBar: AppBar(title: const Text('Settings')),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          const Text('Company name (shown in the header and reports)',
              style: TextStyle(fontWeight: FontWeight.bold)),
          TextField(
              controller: _companyCtrl,
              decoration:
                  const InputDecoration(hintText: 'e.g. Your company name')),
          const SizedBox(height: 20),
          const Text('Company monogram / logo',
              style: TextStyle(fontWeight: FontWeight.bold)),
          const SizedBox(height: 8),
          if (_logo != null)
            Container(
              height: 80,
              alignment: Alignment.centerLeft,
              child: Image.memory(_logo!, fit: BoxFit.contain),
            ),
          Row(
            children: [
              ElevatedButton(
                onPressed: () async {
                  final path = await pickImagePath();
                  if (path == null) return;
                  final bytes = await File(path).readAsBytes();
                  await DbHelper.instance
                      .setSetting('company_logo', base64Encode(bytes));
                  setState(() => _logo = bytes);
                },
                child: const Text('Upload logo'),
              ),
              const SizedBox(width: 8),
              if (_logo != null)
                TextButton(
                  onPressed: () async {
                    await DbHelper.instance.setSetting('company_logo', '');
                    setState(() => _logo = null);
                  },
                  child: const Text('Remove'),
                ),
            ],
          ),
          const SizedBox(height: 24),
          const Text('Day-wise TBT topics (tap a day to edit)',
              style: TextStyle(fontWeight: FontWeight.bold)),
          const SizedBox(height: 8),
          for (var wd = 1; wd <= 7; wd++)
            Card(
              child: ListTile(
                title: Text(
                    '${kWeekdayNames[wd - 1]}: ${_topics[wd]?['title'] ?? ''}'),
                subtitle: Text(_topics[wd]?['detail'] ?? '',
                    maxLines: 2, overflow: TextOverflow.ellipsis),
                trailing: const Icon(Icons.edit),
                onTap: () async {
                  final titleCtrl = TextEditingController(
                      text: _topics[wd]?['title'] ?? '');
                  final detailCtrl = TextEditingController(
                      text: _topics[wd]?['detail'] ?? '');
                  final ok = await showDialog<bool>(
                    context: context,
                    builder: (ctx) => AlertDialog(
                      title: Text('Topic - ${kWeekdayNames[wd - 1]}'),
                      content: SingleChildScrollView(
                        child: Column(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            TextField(
                                controller: titleCtrl,
                                decoration: const InputDecoration(
                                    labelText: 'Topic title')),
                            TextField(
                                controller: detailCtrl,
                                maxLines: 6,
                                decoration: const InputDecoration(
                                    labelText: 'Topic detail')),
                          ],
                        ),
                      ),
                      actions: [
                        TextButton(
                            onPressed: () => Navigator.pop(ctx, false),
                            child: const Text('Cancel')),
                        ElevatedButton(
                            onPressed: () => Navigator.pop(ctx, true),
                            child: const Text('Save')),
                      ],
                    ),
                  );
                  if (ok == true) {
                    setState(() {
                      _topics[wd] = {
                        'title': titleCtrl.text,
                        'detail': detailCtrl.text
                      };
                    });
                  }
                },
              ),
            ),
          const SizedBox(height: 24),
          ElevatedButton(
              onPressed: _save, child: const Text('Save settings')),
        ],
      ),
    );
  }
}

// ---------------- Staff ----------------
class StaffScreen extends StatefulWidget {
  const StaffScreen({super.key});
  @override
  State<StaffScreen> createState() => _StaffScreenState();
}

class _StaffScreenState extends State<StaffScreen> {
  List<Map<String, Object?>> _staff = [];
  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    final list = await DbHelper.instance.staffList();
    if (mounted) setState(() => _staff = list);
  }

  Future<void> _editStaff([Map<String, Object?>? row]) async {
    final name = TextEditingController(text: (row?['name'] as String?) ?? '');
    final iqama =
        TextEditingController(text: (row?['iqama'] as String?) ?? '');
    final nat = TextEditingController(
        text: (row?['nationality'] as String?) ?? '');
    String des = (row?['designation'] as String?) ?? 'Labour';
    String shift = (row?['shift'] as String?) ?? 'Day';
    final ok = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: Text(row == null ? 'Add staff' : 'Edit staff'),
        content: SingleChildScrollView(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              TextField(
                  controller: name,
                  decoration: const InputDecoration(labelText: 'Name')),
              TextField(
                  controller: iqama,
                  decoration: const InputDecoration(labelText: 'Iqama No.')),
              TextField(
                  controller: nat,
                  decoration:
                      const InputDecoration(labelText: 'Nationality')),
              DropdownButtonFormField<String>(
                value: des,
                items: const ['Operator', 'Labour', 'Foreman']
                    .map((e) => DropdownMenuItem(value: e, child: Text(e)))
                    .toList(),
                onChanged: (v) => des = v ?? des,
                decoration: const InputDecoration(labelText: 'Designation'),
              ),
              DropdownButtonFormField<String>(
                value: shift,
                items: const ['Day', 'Night']
                    .map((e) => DropdownMenuItem(value: e, child: Text(e)))
                    .toList(),
                onChanged: (v) => shift = v ?? shift,
                decoration: const InputDecoration(labelText: 'Shift'),
              ),
            ],
          ),
        ),
        actions: [
          TextButton(
              onPressed: () => Navigator.pop(ctx, false),
              child: const Text('Cancel')),
          ElevatedButton(
              onPressed: () => Navigator.pop(ctx, true),
              child: const Text('Save')),
        ],
      ),
    );
    if (ok != true || name.text.trim().isEmpty) return;
    final data = {
      'name': name.text.trim(),
      'iqama': iqama.text.trim(),
      'nationality': nat.text.trim(),
      'designation': des,
      'shift': shift,
    };
    if (row == null) {
      await DbHelper.instance.insertStaff(data);
    } else {
      await DbHelper.instance.updateStaff(row['id'] as int, data);
    }
    _load();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Staff List')),
      floatingActionButton: Column(
        mainAxisSize: MainAxisSize.min,
        children: [
          FloatingActionButton(
            heroTag: 'import',
            mini: true,
            tooltip: 'Photo import',
            onPressed: () async {
              await Navigator.push(
                  context,
                  MaterialPageRoute(
                      builder: (_) => const PhotoImportScreen()));
              _load();
            },
            child: const Icon(Icons.photo_camera),
          ),
          const SizedBox(height: 8),
          FloatingActionButton(
            heroTag: 'add',
            tooltip: 'Add staff',
            onPressed: () => _editStaff(),
            child: const Icon(Icons.add),
          ),
        ],
      ),
      body: _staff.isEmpty
          ? const Center(
              child: Padding(
                padding: EdgeInsets.all(24),
                child: Text(
                  'No staff yet. Tap + to add manually, or use the camera button for photo-assisted import.',
                  textAlign: TextAlign.center,
                ),
              ),
            )
          : ListView.builder(
              itemCount: _staff.length,
              itemBuilder: (ctx, i) {
                final s = _staff[i];
                return ListTile(
                  title: Text((s['name'] as String?) ?? ''),
                  subtitle: Text(
                      '${s['designation']} • ${s['shift']} shift • Iqama: ${s['iqama']} • ${s['nationality']}'),
                  trailing: Row(
                    mainAxisSize: MainAxisSize.min,
                    children: [
                      IconButton(
                          icon: const Icon(Icons.edit),
                          onPressed: () => _editStaff(s)),
                      IconButton(
                        icon: const Icon(Icons.delete),
                        onPressed: () async {
                          await DbHelper.instance
                              .deleteStaff(s['id'] as int);
                          _load();
                        },
                      ),
                    ],
                  ),
                );
              },
            ),
    );
  }
}

// ---------------- Photo-assisted import (review before save) ----------------
class PhotoImportScreen extends StatefulWidget {
  const PhotoImportScreen({super.key});
  @override
  State<PhotoImportScreen> createState() => _PhotoImportScreenState();
}

class _PhotoImportScreenState extends State<PhotoImportScreen> {
  String? _photoPath;
  final List<Map<String, String>> _rows = [];

  Future<void> _addRow() async {
    final name = TextEditingController();
    final iqama = TextEditingController();
    final nat = TextEditingController();
    String des = 'Labour';
    String shift = 'Day';
    final ok = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('Add row from photo'),
        content: SingleChildScrollView(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              TextField(
                  controller: name,
                  decoration: const InputDecoration(labelText: 'Name')),
              TextField(
                  controller: iqama,
                  decoration: const InputDecoration(labelText: 'Iqama No.')),
              TextField(
                  controller: nat,
                  decoration:
                      const InputDecoration(labelText: 'Nationality')),
              DropdownButtonFormField<String>(
                value: des,
                items: const ['Operator', 'Labour', 'Foreman']
                    .map((e) => DropdownMenuItem(value: e, child: Text(e)))
                    .toList(),
                onChanged: (v) => des = v ?? des,
                decoration: const InputDecoration(labelText: 'Designation'),
              ),
              DropdownButtonFormField<String>(
                value: shift,
                items: const ['Day', 'Night']
                    .map((e) => DropdownMenuItem(value: e, child: Text(e)))
                    .toList(),
                onChanged: (v) => shift = v ?? shift,
                decoration: const InputDecoration(labelText: 'Shift'),
              ),
            ],
          ),
        ),
        actions: [
          TextButton(
              onPressed: () => Navigator.pop(ctx, false),
              child: const Text('Cancel')),
          ElevatedButton(
              onPressed: () => Navigator.pop(ctx, true),
              child: const Text('Add to review')),
        ],
      ),
    );
    if (ok == true && name.text.trim().isNotEmpty) {
      setState(() {
        _rows.add({
          'name': name.text.trim(),
          'iqama': iqama.text.trim(),
          'nationality': nat.text.trim(),
          'designation': des,
          'shift': shift,
        });
      });
    }
  }

  Future<void> _saveAll() async {
    final existing = await DbHelper.instance.staffList();
    var added = 0;
    for (final r in _rows) {
      final dup = existing.any((e) =>
          (r['iqama']!.isNotEmpty && e['iqama'] == r['iqama']) ||
          (e['name'] == r['name']));
      if (dup) continue;
      await DbHelper.instance.insertStaff({
        'name': r['name']!,
        'iqama': r['iqama']!,
        'nationality': r['nationality']!,
        'designation': r['designation']!,
        'shift': r['shift']!,
      });
      added++;
    }
    if (mounted) showMsg(context, 'Saved $added staff (duplicates skipped)');
    if (mounted) Navigator.pop(context);
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Photo Import - Review')),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          const Text(
            'Pick a photo of a staff list, read each row from the photo, add it below and review before saving. Automatic text recognition is not included in this test build.',
            style: TextStyle(fontSize: 13),
          ),
          const SizedBox(height: 12),
          ElevatedButton(
            onPressed: () async {
              final path = await pickImagePath();
              if (path != null) setState(() => _photoPath = path);
            },
            child: const Text('Pick staff list photo'),
          ),
          if (_photoPath != null) ...[
            const SizedBox(height: 12),
            Image.file(File(_photoPath!), height: 220, fit: BoxFit.contain),
          ],
          const SizedBox(height: 16),
          Row(
            children: [
              ElevatedButton(
                  onPressed: _addRow, child: const Text('Add row')),
              const SizedBox(width: 8),
              ElevatedButton(
                  onPressed: _rows.isEmpty ? null : _saveAll,
                  child: const Text('Save all reviewed rows')),
            ],
          ),
          const SizedBox(height: 12),
          for (var i = 0; i < _rows.length; i++)
            ListTile(
              title: Text(_rows[i]['name']!),
              subtitle: Text(
                  '${_rows[i]['designation']} • ${_rows[i]['shift']} • Iqama: ${_rows[i]['iqama']} • ${_rows[i]['nationality']}'),
              trailing: IconButton(
                icon: const Icon(Icons.close),
                onPressed: () => setState(() => _rows.removeAt(i)),
              ),
            ),
        ],
      ),
    );
  }
}

// ---------------- Toolbox Talk ----------------
class AttendeeDraft {
  String name;
  String iqama;
  String designation;
  String nationality;
  String shift;
  String remarks;
  Uint8List? signature;
  AttendeeDraft(
      {required this.name,
      required this.iqama,
      required this.designation,
      required this.nationality,
      required this.shift,
      this.remarks = '',
      this.signature});
}

class TbtScreen extends StatefulWidget {
  const TbtScreen({super.key});
  @override
  State<TbtScreen> createState() => _TbtScreenState();
}

class _TbtScreenState extends State<TbtScreen> {
  bool _showForm = false;
  List<Map<String, Object?>> _saved = [];
  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    final d = await DbHelper.instance.db;
    final rows = await d.query('tbt', orderBy: 'date DESC, id DESC');
    if (mounted) setState(() => _saved = rows);
  }

  @override
  Widget build(BuildContext context) {
    if (_showForm) {
      return _TbtForm(onSaved: () {
        setState(() => _showForm = false);
        _load();
      });
    }
    return Scaffold(
      appBar: AppBar(title: const Text('Toolbox Talk')),
      floatingActionButton: FloatingActionButton(
        onPressed: () => setState(() => _showForm = true),
        child: const Icon(Icons.add),
      ),
      body: _saved.isEmpty
          ? const Center(child: Text('No Toolbox Talk records yet. Tap + to create one.'))
          : ListView.builder(
              itemCount: _saved.length,
              itemBuilder: (ctx, i) {
                final t = _saved[i];
                return ListTile(
                  title: Text(
                      '${t['date']} • ${t['topic']}'),
                  subtitle: Text(
                      '${t['terminal']} • ${t['shift']} shift • ${t['project']}'),
                  trailing: const Icon(Icons.chevron_right),
                  onTap: () => Navigator.push(
                      context,
                      MaterialPageRoute(
                          builder: (_) =>
                              TbtViewScreen(tbtId: t['id'] as int))),
                );
              },
            ),
    );
  }
}

class _TbtForm extends StatefulWidget {
  final VoidCallback onSaved;
  const _TbtForm({required this.onSaved});
  @override
  State<_TbtForm> createState() => _TbtFormState();
}

class _TbtFormState extends State<_TbtForm> {
  DateTime _date = DateTime.now();
  final _project = TextEditingController(
      text: 'Stripping and Stuffing Operation');
  final _contractor = TextEditingController();
  final _terminal = TextEditingController();
  final _location =
      TextEditingController(text: 'King Abdul Aziz Port Dammam');
  String _shift = 'Day';
  String _topic = '';
  final _topicDetail = TextEditingController();
  String? _photoPath;
  List<Map<String, Object?>> _staff = [];
  final List<AttendeeDraft> _attendees = [];

  @override
  void initState() {
    super.initState();
    _init();
  }

  Future<void> _init() async {
    final topics = await DbHelper.instance.topics();
    final staff = await DbHelper.instance.staffList();
    final t = topics[_date.weekday] ?? kDefaultTopics[_date.weekday]!;
    if (mounted) {
      setState(() {
        _topic = t['title']!;
        _topicDetail.text = t['detail']!;
        _staff = staff;
      });
    }
  }

  Future<void> _pickDate() async {
    final d = await showDatePicker(
        context: context,
        initialDate: _date,
        firstDate: DateTime(2020),
        lastDate: DateTime(2100));
    if (d == null) return;
    final topics = await DbHelper.instance.topics();
    final t = topics[d.weekday] ?? kDefaultTopics[d.weekday]!;
    setState(() {
      _date = d;
      _topic = t['title']!;
      _topicDetail.text = t['detail']!;
    });
  }

  Future<void> _save() async {
    if (_topic.trim().isEmpty) {
      showMsg(context, 'Topic is required');
      return;
    }
    final d = await DbHelper.instance.db;
    final id = await d.insert('tbt', {
      'date': fmtDate(_date),
      'project': _project.text.trim(),
      'contractor': _contractor.text.trim(),
      'terminal': _terminal.text.trim(),
      'shift': _shift,
      'location': _location.text.trim(),
      'topic': _topic.trim(),
      'topic_detail': _topicDetail.text.trim(),
      'photo_path': _photoPath,
    });
    for (final a in _attendees) {
      await d.insert('tbt_attendees', {
        'tbt_id': id,
        'name': a.name,
        'iqama': a.iqama,
        'designation': a.designation,
        'nationality': a.nationality,
        'shift': a.shift,
        'remarks': a.remarks,
        'signature': a.signature,
      });
    }
    if (mounted) showMsg(context, 'Toolbox Talk saved');
    widget.onSaved();
  }

  @override
  void dispose() {
    _project.dispose();
    _contractor.dispose();
    _terminal.dispose();
    _location.dispose();
    _topicDetail.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('New Toolbox Talk')),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          ListTile(
            title: Text('Date: ${fmtDate(_date)}'),
            trailing: const Icon(Icons.calendar_today),
            onTap: _pickDate,
          ),
          TextField(
              controller: _project,
              decoration: const InputDecoration(labelText: 'Project')),
          TextField(
              controller: _contractor,
              decoration: const InputDecoration(labelText: 'Contractor')),
          TextField(
              controller: _terminal,
              decoration: const InputDecoration(labelText: 'Terminal')),
          TextField(
              controller: _location,
              decoration: const InputDecoration(labelText: 'Location')),
          DropdownButtonFormField<String>(
            value: _shift,
            items: const ['Day', 'Night']
                .map((e) => DropdownMenuItem(value: e, child: Text(e)))
                .toList(),
            onChanged: (v) => setState(() => _shift = v ?? _shift),
            decoration: const InputDecoration(labelText: 'Duty Shift'),
          ),
          const SizedBox(height: 12),
          Text('Topic (${kWeekdayNames[_date.weekday - 1]}): $_topic',
              style: const TextStyle(fontWeight: FontWeight.bold)),
          TextField(
              controller: _topicDetail,
              maxLines: 6,
              decoration:
                  const InputDecoration(labelText: 'Topic detail')),
          const SizedBox(height: 12),
          Row(
            children: [
              ElevatedButton(
                onPressed: () async {
                  final path = await pickImagePath();
                  if (path != null) setState(() => _photoPath = path);
                },
                child: const Text('Attach photo (optional)'),
              ),
              if (_photoPath != null) ...[
                const SizedBox(width: 8),
                const Icon(Icons.check_circle, color: Colors.green),
              ],
            ],
          ),
          const SizedBox(height: 20),
          const Text('Attendees (tap staff to add, then sign)',
              style: TextStyle(fontWeight: FontWeight.bold)),
          Wrap(
            spacing: 6,
            children: _staff.map((s) {
              final name = (s['name'] as String?) ?? '';
              final added = _attendees.any((a) => a.name == name);
              return FilterChip(
                label: Text(name),
                selected: added,
                onSelected: (v) {
                  setState(() {
                    if (v) {
                      _attendees.add(AttendeeDraft(
                        name: name,
                        iqama: (s['iqama'] as String?) ?? '',
                        designation: (s['designation'] as String?) ?? '',
                        nationality: (s['nationality'] as String?) ?? '',
                        shift: (s['shift'] as String?) ?? 'Day',
                      ));
                    } else {
                      _attendees.removeWhere((a) => a.name == name);
                    }
                  });
                },
              );
            }).toList(),
          ),
          const SizedBox(height: 8),
          TextButton(
            onPressed: () async {
              final name = TextEditingController();
              final iqama = TextEditingController();
              final ok = await showDialog<bool>(
                context: context,
                builder: (ctx) => AlertDialog(
                  title: const Text('Add attendee manually'),
                  content: Column(
                    mainAxisSize: MainAxisSize.min,
                    children: [
                      TextField(
                          controller: name,
                          decoration:
                              const InputDecoration(labelText: 'Name')),
                      TextField(
                          controller: iqama,
                          decoration: const InputDecoration(
                              labelText: 'Iqama No.')),
                    ],
                  ),
                  actions: [
                    TextButton(
                        onPressed: () => Navigator.pop(ctx, false),
                        child: const Text('Cancel')),
                    ElevatedButton(
                        onPressed: () => Navigator.pop(ctx, true),
                        child: const Text('Add')),
                  ],
                ),
              );
              if (ok == true && name.text.trim().isNotEmpty) {
                setState(() {
                  _attendees.add(AttendeeDraft(
                    name: name.text.trim(),
                    iqama: iqama.text.trim(),
                    designation: '',
                    nationality: '',
                    shift: _shift,
                  ));
                });
              }
            },
            child: const Text('+ Add attendee manually'),
          ),
          for (var i = 0; i < _attendees.length; i++)
            ListTile(
              title: Text(_attendees[i].name),
              subtitle: Text(
                  '${_attendees[i].designation} • Iqama: ${_attendees[i].iqama}'),
              trailing: Row(
                mainAxisSize: MainAxisSize.min,
                children: [
                  if (_attendees[i].signature != null)
                    const Icon(Icons.check, color: Colors.green),
                  TextButton(
                    onPressed: () async {
                      final sig = await askSignature(context);
                      if (sig != null) {
                        setState(() => _attendees[i].signature = sig);
                      }
                    },
                    child: const Text('Sign'),
                  ),
                ],
              ),
            ),
          const SizedBox(height: 20),
          ElevatedButton(onPressed: _save, child: const Text('Save TBT')),
        ],
      ),
    );
  }
}

class TbtViewScreen extends StatelessWidget {
  final int tbtId;
  const TbtViewScreen({super.key, required this.tbtId});
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Toolbox Talk Record')),
      body: FutureBuilder(
        future: () async {
          final d = await DbHelper.instance.db;
          final t = await d.query('tbt', where: 'id = ?', whereArgs: [tbtId]);
          final a = await d.query('tbt_attendees',
              where: 'tbt_id = ?', whereArgs: [tbtId]);
          final company = await DbHelper.instance.companyName();
          final logo = await DbHelper.instance.logoBytes();
          return {'t': t.first, 'a': a, 'company': company, 'logo': logo};
        }(),
        builder: (ctx, snap) {
          if (!snap.hasData) {
            return const Center(child: CircularProgressIndicator());
          }
          final data = snap.data!;
          final t = data['t'] as Map<String, Object?>;
          final atts = data['a'] as List<Map<String, Object?>>;
          final logo = data['logo'] as Uint8List?;
          return ListView(
            padding: const EdgeInsets.all(16),
            children: [
              if (logo != null)
                SizedBox(
                    height: 70,
                    child: Image.memory(logo, fit: BoxFit.contain)),
              Text((data['company'] as String?) ?? '',
                  style: const TextStyle(
                      fontSize: 20, fontWeight: FontWeight.bold)),
              const SizedBox(height: 8),
              Text('Project: ${t['project']}'),
              Text('Date: ${t['date']}'),
              Text('Contractor: ${t['contractor']}'),
              Text('Terminal: ${t['terminal']}'),
              Text('Duty Shift: ${t['shift']}'),
              Text('Location: ${t['location']}'),
              const SizedBox(height: 12),
              Text('Topic: ${t['topic']}',
                  style: const TextStyle(fontWeight: FontWeight.bold)),
              Text('${t['topic_detail']}'),
              if (t['photo_path'] != null &&
                  File(t['photo_path'] as String).existsSync())
                Padding(
                  padding: const EdgeInsets.only(top: 12),
                  child: Image.file(File(t['photo_path'] as String),
                      height: 200, fit: BoxFit.contain),
                ),
              const SizedBox(height: 16),
              const Text('Attendees',
                  style: TextStyle(fontWeight: FontWeight.bold)),
              const SizedBox(height: 4),
              SingleChildScrollView(
                scrollDirection: Axis.horizontal,
                child: DataTable(
                  columns: const [
                    DataColumn(label: Text('S.No')),
                    DataColumn(label: Text('Iqama No.')),
                    DataColumn(label: Text('Name')),
                    DataColumn(label: Text('Designation')),
                    DataColumn(label: Text('Nationality')),
                    DataColumn(label: Text('Shift')),
                    DataColumn(label: Text('Signature')),
                    DataColumn(label: Text('Remarks')),
                  ],
                  rows: [
                    for (var i = 0; i < atts.length; i++)
                      DataRow(cells: [
                        DataCell(Text('${i + 1}')),
                        DataCell(Text((atts[i]['iqama'] as String?) ?? '')),
                        DataCell(Text((atts[i]['name'] as String?) ?? '')),
                        DataCell(
                            Text((atts[i]['designation'] as String?) ?? '')),
                        DataCell(
                            Text((atts[i]['nationality'] as String?) ?? '')),
                        DataCell(Text((atts[i]['shift'] as String?) ?? '')),
                        DataCell(atts[i]['signature'] != null
                            ? Image.memory(
                                atts[i]['signature'] as Uint8List,
                                width: 60,
                                height: 28,
                                fit: BoxFit.contain)
                            : const Text('-')),
                        DataCell(Text((atts[i]['remarks'] as String?) ?? '')),
                      ]),
                  ],
                ),
              ),
            ],
          );
        },
      ),
    );
  }
}

// ---------------- Attendance ----------------
class AttendanceScreen extends StatefulWidget {
  const AttendanceScreen({super.key});
  @override
  State<AttendanceScreen> createState() => _AttendanceScreenState();
}

class _AttendanceScreenState extends State<AttendanceScreen> {
  DateTime _date = DateTime.now();
  List<Map<String, Object?>> _staff = [];
  final Map<int, String> _shifts = {};
  final Map<int, String> _remarks = {};
  final Map<int, String> _ot = {};
  final Map<int, Uint8List?> _sigs = {};
  bool _showSummary = false;

  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    final staff = await DbHelper.instance.staffList();
    final d = await DbHelper.instance.db;
    final rows = await d.query('attendance',
        where: 'date = ?', whereArgs: [fmtDate(_date)]);
    if (mounted) {
      setState(() {
        _staff = staff;
        _shifts.clear();
        _remarks.clear();
        _ot.clear();
        _sigs.clear();
        for (final r in rows) {
          final sid = r['staff_id'] as int;
          _shifts[sid] = (r['shift'] as String?) ?? 'Day';
          _remarks[sid] = (r['remarks'] as String?) ?? '';
          _ot[sid] = (r['ot_hours'] as num?)?.toString() ?? '';
          _sigs[sid] = r['signature'] as Uint8List?;
        }
      });
    }
  }

  Future<void> _save() async {
    final d = await DbHelper.instance.db;
    await d.delete('attendance',
        where: 'date = ?', whereArgs: [fmtDate(_date)]);
    for (final s in _staff) {
      final sid = s['id'] as int;
      await d.insert('attendance', {
        'date': fmtDate(_date),
        'staff_id': sid,
        'staff_name': s['name'],
        'designation': s['designation'],
        'shift': _shifts[sid] ?? (s['shift'] as String?) ?? 'Day',
        'remarks': _remarks[sid] ?? '',
        'ot_hours': double.tryParse(_ot[sid] ?? '') ?? 0,
        'signature': _sigs[sid],
      });
    }
    if (mounted) showMsg(context, 'Attendance saved for ${fmtDate(_date)}');
  }

  @override
  Widget build(BuildContext context) {
    if (_showSummary) return _OtSummary(onBack: () {
      setState(() => _showSummary = false);
    });
    return Scaffold(
      appBar: AppBar(
        title: const Text('Attendance'),
        actions: [
          TextButton(
            onPressed: () => setState(() => _showSummary = true),
            child: const Text('OT summary',
                style: TextStyle(color: Colors.white)),
          ),
        ],
      ),
      body: Column(
        children: [
          ListTile(
            title: Text('Date: ${fmtDate(_date)}'),
            trailing: const Icon(Icons.calendar_today),
            onTap: () async {
              final d = await showDatePicker(
                  context: context,
                  initialDate: _date,
                  firstDate: DateTime(2020),
                  lastDate: DateTime(2100));
              if (d != null) {
                setState(() => _date = d);
                _load();
              }
            },
          ),
          Expanded(
            child: _staff.isEmpty
                ? const Center(
                    child: Text('No staff yet. Add staff first.'))
                : ListView.builder(
                    itemCount: _staff.length,
                    itemBuilder: (ctx, i) {
                      final s = _staff[i];
                      final sid = s['id'] as int;
                      return Card(
                        margin: const EdgeInsets.symmetric(
                            horizontal: 12, vertical: 4),
                        child: Padding(
                          padding: const EdgeInsets.all(8),
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text((s['name'] as String?) ?? '',
                                  style: const TextStyle(
                                      fontWeight: FontWeight.bold)),
                              Text(
                                  '${s['designation']} • ${(s['shift'] as String?) ?? ''} shift'),
                              Row(
                                children: [
                                  Expanded(
                                    child: DropdownButtonFormField<String>(
                                      value: _shifts[sid] ??
                                          (s['shift'] as String?) ??
                                          'Day',
                                      items: const ['Day', 'Night']
                                          .map((e) => DropdownMenuItem(
                                              value: e, child: Text(e)))
                                          .toList(),
                                      onChanged: (v) => setState(
                                          () => _shifts[sid] = v ?? 'Day'),
                                      decoration: const InputDecoration(
                                          labelText: 'Shift'),
                                    ),
                                  ),
                                  const SizedBox(width: 8),
                                  Expanded(
                                    child: TextFormField(
                                      initialValue: _ot[sid] ?? '',
                                      keyboardType:
                                          const TextInputType.numberWithOptions(
                                              decimal: true),
                                      decoration: const InputDecoration(
                                          labelText: 'OT hours (manual)'),
                                      onChanged: (v) => _ot[sid] = v,
                                    ),
                                  ),
                                ],
                              ),
                              Row(
                                children: [
                                  Expanded(
                                    child: TextFormField(
                                      initialValue: _remarks[sid] ?? '',
                                      decoration: const InputDecoration(
                                          labelText: 'Remarks'),
                                      onChanged: (v) => _remarks[sid] = v,
                                    ),
                                  ),
                                  TextButton(
                                    onPressed: () async {
                                      final sig =
                                          await askSignature(context);
                                      if (sig != null) {
                                        setState(
                                            () => _sigs[sid] = sig);
                                      }
                                    },
                                    child: Text(_sigs[sid] != null
                                        ? 'Signed ✓'
                                        : 'Sign'),
                                  ),
                                ],
                              ),
                            ],
                          ),
                        ),
                      );
                    },
                  ),
          ),
          Padding(
            padding: const EdgeInsets.all(12),
            child: ElevatedButton(
                onPressed: _save,
                child: const Text('Save attendance')),
          ),
        ],
      ),
    );
  }
}

class _OtSummary extends StatelessWidget {
  final VoidCallback onBack;
  const _OtSummary({required this.onBack});
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Overtime Summary'),
        leading: IconButton(
            icon: const Icon(Icons.arrow_back), onPressed: onBack),
      ),
      body: FutureBuilder(
        future: () async {
          final d = await DbHelper.instance.db;
          final rows = await d.query('attendance',
              orderBy: 'staff_name ASC, date ASC');
          final totals = <String, double>{};
          final des = <String, String>{};
          for (final r in rows) {
            final name = (r['staff_name'] as String?) ?? '';
            totals[name] =
                (totals[name] ?? 0) + ((r['ot_hours'] as num?) ?? 0).toDouble();
            des[name] = (r['designation'] as String?) ?? '';
          }
          return {'rows': rows, 'totals': totals, 'des': des};
        }(),
        builder: (ctx, snap) {
          if (!snap.hasData) {
            return const Center(child: CircularProgressIndicator());
          }
          final totals = snap.data!['totals'] as Map<String, double>;
          final des = snap.data!['des'] as Map<String, String>;
          final rows = snap.data!['rows'] as List<Map<String, Object?>>;
          if (totals.isEmpty) {
            return const Center(child: Text('No attendance records yet.'));
          }
          return ListView(
            padding: const EdgeInsets.all(16),
            children: [
              const Text('Total OT hours per worker (manual entries only)',
                  style: TextStyle(fontWeight: FontWeight.bold)),
              const SizedBox(height: 8),
              for (final e in totals.entries)
                ListTile(
                  title: Text(e.key),
                  subtitle: Text(des[e.key] ?? ''),
                  trailing: Text('${e.value} hrs',
                      style: const TextStyle(fontWeight: FontWeight.bold)),
                ),
              const Divider(),
              const Text('Date-wise detail',
                  style: TextStyle(fontWeight: FontWeight.bold)),
              for (final r in rows)
                ListTile(
                  title: Text(
                      '${r['staff_name']} — ${r['date']} (${r['shift']})'),
                  trailing: Text('${r['ot_hours']} hrs'),
                ),
            ],
          );
        },
      ),
    );
  }
}

// ---------------- Container Moves ----------------
class MovesScreen extends StatefulWidget {
  const MovesScreen({super.key});
  @override
  State<MovesScreen> createState() => _MovesScreenState();
}

class _MovesScreenState extends State<MovesScreen> {
  DateTime _date = DateTime.now();
  List<Map<String, Object?>> _moves = [];
  List<Map<String, Object?>> _staff = [];
  final _container = TextEditingController();
  String _shift = 'Day';
  final _foreman = TextEditingController();
  final _search = TextEditingController();
  final Set<int> _crew = {};
  // lookup
  final _lookupName = TextEditingController();
  DateTime _lookupDate = DateTime.now();
  List<Map<String, Object?>> _lookupMoves = [];
  int _lookupTotal = 0;

  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    final d = await DbHelper.instance.db;
    final moves = await d.query('moves',
        where: 'date = ?', whereArgs: [fmtDate(_date)], orderBy: 'id DESC');
    final staff = await DbHelper.instance.staffList();
    if (mounted) setState(() {
      _moves = moves;
      _staff = staff;
    });
  }

  Future<Map<String, int>> _countsForDate(String date) async {
    final d = await DbHelper.instance.db;
    final rows = await d.rawQuery(
        'SELECT mc.staff_name as name, COUNT(*) as c FROM move_crew mc JOIN moves m ON mc.move_id = m.id WHERE m.date = ? GROUP BY mc.staff_name',
        [date]);
    final out = <String, int>{};
    for (final r in rows) {
      out[(r['name'] as String?) ?? ''] = (r['c'] as int?) ?? 0;
    }
    return out;
  }

  Future<void> _saveMove() async {
    if (_container.text.trim().isEmpty) {
      showMsg(context, 'Container number is required');
      return;
    }
    final d = await DbHelper.instance.db;
    final id = await d.insert('moves', {
      'date': fmtDate(_date),
      'container': _container.text.trim(),
      'shift': _shift,
      'foreman': _foreman.text.trim(),
    });
    for (final sid in _crew) {
      final s = _staff.firstWhere((e) => e['id'] == sid, orElse: () => {});
      await d.insert('move_crew', {
        'move_id': id,
        'staff_id': sid,
        'staff_name': s['name'] ?? '',
      });
    }
    _container.clear();
    _crew.clear();
    _load();
    if (mounted) showMsg(context, 'Container move saved');
  }

  Future<void> _runLookup() async {
    final name = _lookupName.text.trim();
    if (name.isEmpty) return;
    final d = await DbHelper.instance.db;
    final rows = await d.rawQuery(
        'SELECT m.container as container, m.date as date FROM move_crew mc JOIN moves m ON mc.move_id = m.id WHERE m.date = ? AND mc.staff_name LIKE ? ORDER BY m.id',
        [fmtDate(_lookupDate), '%$name%']);
    if (mounted) {
      setState(() {
        _lookupMoves = rows;
        _lookupTotal = rows.length;
      });
    }
  }

  @override
  void dispose() {
    _container.dispose();
    _search.dispose();
    _foreman.dispose();
    _lookupName.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final filtered = _staff.where((s) {
      final q = _search.text.trim().toLowerCase();
      if (q.isEmpty) return true;
      return ((s['name'] as String?) ?? '').toLowerCase().contains(q);
    }).toList();
    return Scaffold(
      appBar: AppBar(title: const Text('Container Moves')),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          ListTile(
            title: Text('Date: ${fmtDate(_date)}'),
            trailing: const Icon(Icons.calendar_today),
            onTap: () async {
              final d = await showDatePicker(
                  context: context,
                  initialDate: _date,
                  firstDate: DateTime(2020),
                  lastDate: DateTime(2100));
              if (d != null) {
                setState(() => _date = d);
                _load();
              }
            },
          ),
          TextField(
              controller: _container,
              decoration:
                  const InputDecoration(labelText: 'Container number')),
          TextField(
              controller: _foreman,
              decoration: const InputDecoration(labelText: 'Foreman')),
          DropdownButtonFormField<String>(
            value: _shift,
            items: const ['Day', 'Night']
                .map((e) => DropdownMenuItem(value: e, child: Text(e)))
                .toList(),
            onChanged: (v) => setState(() => _shift = v ?? _shift),
            decoration: const InputDecoration(labelText: 'Shift'),
          ),
          const SizedBox(height: 12),
          const Text('Crew - search a name and tap to add/remove',
              style: TextStyle(fontWeight: FontWeight.bold)),
          TextField(
            controller: _search,
            decoration: const InputDecoration(hintText: 'Search staff name'),
            onChanged: (_) => setState(() {}),
          ),
          Wrap(
            spacing: 6,
            children: filtered.map((s) {
              final sid = s['id'] as int;
              return FilterChip(
                label: Text((s['name'] as String?) ?? ''),
                selected: _crew.contains(sid),
                onSelected: (v) {
                  setState(() {
                    if (v) {
                      _crew.add(sid);
                    } else {
                      _crew.remove(sid);
                    }
                  });
                },
              );
            }).toList(),
          ),
          const SizedBox(height: 12),
          ElevatedButton(
              onPressed: _saveMove, child: const Text('Save container move')),
          const SizedBox(height: 20),
          const Text('Moves on this date',
              style: TextStyle(fontWeight: FontWeight.bold)),
          for (final m in _moves)
            FutureBuilder<List<Map<String, Object?>>>(
              future: () async {
                final d = await DbHelper.instance.db;
                return d.query('move_crew',
                    where: 'move_id = ?', whereArgs: [m['id']]);
              }(),
              builder: (ctx, snap) {
                final crew = snap.data ?? [];
                final names =
                    crew.map((c) => c['staff_name'] as String?).join(', ');
                return ListTile(
                  title: Text('${m['container']} • ${m['shift']} shift'),
                  subtitle: Text(
                      'Foreman: ${m['foreman']} • Crew: $names'),
                );
              },
            ),
          const SizedBox(height: 12),
          FutureBuilder<Map<String, int>>(
            future: _countsForDate(fmtDate(_date)),
            builder: (ctx, snap) {
              final counts = snap.data ?? {};
              var total = 0;
              counts.values.forEach((v) => total += v);
              return Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text('Daily summary - total moves: $total',
                      style: const TextStyle(fontWeight: FontWeight.bold)),
                  for (final e in counts.entries)
                    Text('${e.key}: ${e.value} moves'),
                ],
              );
            },
          ),
          const Divider(height: 32),
          const Text('Date-wise move lookup',
              style: TextStyle(fontWeight: FontWeight.bold)),
          ListTile(
            title: Text('Lookup date: ${fmtDate(_lookupDate)}'),
            trailing: const Icon(Icons.calendar_today),
            onTap: () async {
              final d = await showDatePicker(
                  context: context,
                  initialDate: _lookupDate,
                  firstDate: DateTime(2020),
                  lastDate: DateTime(2100));
              if (d != null) setState(() => _lookupDate = d);
            },
          ),
          TextField(
              controller: _lookupName,
              decoration:
                  const InputDecoration(labelText: 'Labour name')),
          const SizedBox(height: 8),
          ElevatedButton(
              onPressed: _runLookup, child: const Text('Search moves')),
          if (_lookupMoves.isNotEmpty) ...[
            const SizedBox(height: 8),
            Text('Total moves for this worker on ${fmtDate(_lookupDate)}: $_lookupTotal',
                style: const TextStyle(fontWeight: FontWeight.bold)),
            for (final r in _lookupMoves)
              Text('• ${r['container']}'),
          ],
        ],
      ),
    );
  }
}

// ---------------- Damage Report ----------------
class DamageScreen extends StatefulWidget {
  const DamageScreen({super.key});
  @override
  State<DamageScreen> createState() => _DamageScreenState();
}

class _DamageScreenState extends State<DamageScreen> {
  final _container = TextEditingController();
  final _search = TextEditingController();
  DateTime _date = DateTime.now();
  String _shift = 'Day';
  final _foreman = TextEditingController();
  final _labour = TextEditingController();
  String _origin = 'Original';
  String _clientReport = 'Not made';
  DateTime? _clientDate;
  String? _photoPath;
  final _notes = TextEditingController();
  List<Map<String, Object?>> _list = [];

  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    final d = await DbHelper.instance.db;
    final rows = await d.query('damage', orderBy: 'date DESC, id DESC');
    if (mounted) setState(() => _list = rows);
  }

  Future<void> _autofillFromMoves() async {
    final c = _container.text.trim();
    if (c.isEmpty) return;
    final d = await DbHelper.instance.db;
    final moves = await d.query('moves',
        where: 'container = ?', whereArgs: [c], orderBy: 'id DESC', limit: 1);
    if (moves.isEmpty) return;
    final move = moves.first;
    final crew = await d.query('move_crew',
        where: 'move_id = ?', whereArgs: [move['id']]);
    final names =
        crew.map((e) => e['staff_name'] as String?).join(', ');
    setState(() {
      _shift = (move['shift'] as String?) ?? _shift;
      _foreman.text = (move['foreman'] as String?) ?? '';
      _labour.text = names;
    });
  }

  Future<void> _save() async {
    if (_container.text.trim().isEmpty) {
      showMsg(context, 'Container number is required');
      return;
    }
    final d = await DbHelper.instance.db;
    await d.insert('damage', {
      'container': _container.text.trim(),
      'date': fmtDate(_date),
      'shift': _shift,
      'foreman': _foreman.text.trim(),
      'labour': _labour.text.trim(),
      'origin': _origin,
      'client_report': _clientReport,
      'client_report_date':
          _clientDate != null ? fmtDate(_clientDate!) : '',
      'photo_path': _photoPath,
      'notes': _notes.text.trim(),
    });
    _container.clear();
    _foreman.clear();
    _labour.clear();
    _notes.clear();
    _photoPath = null;
    _load();
    if (mounted) showMsg(context, 'Damage report saved');
  }

  @override
  void dispose() {
    _container.dispose();
    _search.dispose();
    _foreman.dispose();
    _labour.dispose();
    _notes.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final q = _search.text.trim().toLowerCase();
    final filtered = _list.where((r) {
      if (q.isEmpty) return true;
      return ((r['container'] as String?) ?? '').toLowerCase().contains(q);
    }).toList();
    return Scaffold(
      appBar: AppBar(title: const Text('Damage Report')),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          TextField(
            controller: _container,
            decoration:
                const InputDecoration(labelText: 'Container number'),
            onChanged: (_) => _autofillFromMoves(),
          ),
          const Text(
            'Shift, foreman and labour auto-fill from Container Moves if this container was recorded there; otherwise enter manually.',
            style: TextStyle(fontSize: 12, color: Colors.grey),
          ),
          ListTile(
            title: Text('Date: ${fmtDate(_date)}'),
            trailing: const Icon(Icons.calendar_today),
            onTap: () async {
              final d = await showDatePicker(
                  context: context,
                  initialDate: _date,
                  firstDate: DateTime(2020),
                  lastDate: DateTime(2100));
              if (d != null) setState(() => _date = d);
            },
          ),
          DropdownButtonFormField<String>(
            value: _shift,
            items: const ['Day', 'Night']
                .map((e) => DropdownMenuItem(value: e, child: Text(e)))
                .toList(),
            onChanged: (v) => setState(() => _shift = v ?? _shift),
            decoration: const InputDecoration(labelText: 'Shift'),
          ),
          TextField(
              controller: _foreman,
              decoration: const InputDecoration(labelText: 'Foreman')),
          TextField(
              controller: _labour,
              decoration: const InputDecoration(labelText: 'Labour')),
          DropdownButtonFormField<String>(
            value: _origin,
            items: const ['Original', 'Caused by labour']
                .map((e) => DropdownMenuItem(value: e, child: Text(e)))
                .toList(),
            onChanged: (v) => setState(() => _origin = v ?? _origin),
            decoration:
                const InputDecoration(labelText: 'Damage origin'),
          ),
          DropdownButtonFormField<String>(
            value: _clientReport,
            items: const ['Not made', 'Made']
                .map((e) => DropdownMenuItem(value: e, child: Text(e)))
                .toList(),
            onChanged: (v) => setState(() => _clientReport = v ?? _clientReport),
            decoration: const InputDecoration(
                labelText: 'Client damage report made?'),
          ),
          if (_clientReport == 'Made')
            ListTile(
              title: Text(_clientDate == null
                  ? 'Client report date: not set'
                  : 'Client report date: ${fmtDate(_clientDate!)}'),
              trailing: const Icon(Icons.calendar_today),
              onTap: () async {
                final d = await showDatePicker(
                    context: context,
                    initialDate: DateTime.now(),
                    firstDate: DateTime(2020),
                    lastDate: DateTime(2100));
                if (d != null) setState(() => _clientDate = d);
              },
            ),
          Row(
            children: [
              ElevatedButton(
                onPressed: () async {
                  final path = await pickImagePath();
                  if (path != null) setState(() => _photoPath = path);
                },
                child: const Text('Attach photo'),
              ),
              if (_photoPath != null) ...[
                const SizedBox(width: 8),
                const Icon(Icons.check_circle, color: Colors.green),
              ],
            ],
          ),
          TextField(
              controller: _notes,
              maxLines: 3,
              decoration: const InputDecoration(labelText: 'Details / notes')),
          const SizedBox(height: 12),
          ElevatedButton(
              onPressed: _save, child: const Text('Save damage report')),
          const Divider(height: 32),
          TextField(
            controller: _search,
            decoration: const InputDecoration(
                hintText: 'Search by container number'),
            onChanged: (_) => setState(() {}),
          ),
          for (final r in filtered)
            ListTile(
              title: Text(
                  '${r['container']} • ${r['date']} • ${r['origin']}'),
              subtitle: Text(
                  'Foreman: ${r['foreman']} • Labour: ${r['labour']} • Client report: ${r['client_report']}${(r['client_report_date'] as String?)!.isNotEmpty ? ' (${r['client_report_date']})' : ''}'),
              trailing: r['photo_path'] != null &&
                      File(r['photo_path'] as String).existsSync()
                  ? const Icon(Icons.photo)
                  : null,
            ),
        ],
      ),
    );
  }
}

// ---------------- Export ----------------
class ExportScreen extends StatefulWidget {
  const ExportScreen({super.key});
  @override
  State<ExportScreen> createState() => _ExportScreenState();
}

class _ExportScreenState extends State<ExportScreen> {
  DateTime _from = DateTime.now().subtract(const Duration(days: 6));
  DateTime _to = DateTime.now();
  bool _busy = false;

  Future<List<Map<String, Object?>>> _tbtsInRange() async {
    final d = await DbHelper.instance.db;
    return d.query('tbt',
        where: 'date >= ? AND date <= ?',
        whereArgs: [fmtDate(_from), fmtDate(_to)],
        orderBy: 'date ASC, id ASC');
  }

  Future<Map<int, List<Map<String, Object?>>>> _attendeesFor(
      List<int> ids) async {
    final d = await DbHelper.instance.db;
    final out = <int, List<Map<String, Object?>>>{};
    for (final id in ids) {
      out[id] =
          await d.query('tbt_attendees', where: 'tbt_id = ?', whereArgs: [id]);
    }
    return out;
  }

  Future<void> _exportPdf() async {
    setState(() => _busy = true);
    try {
      final tbts = await _tbtsInRange();
      if (tbts.isEmpty) {
        showMsg(context, 'No TBT records in this range');
        return;
      }
      final ids = tbts.map((t) => t['id'] as int).toList();
      final atts = await _attendeesFor(ids);
      final company = await DbHelper.instance.companyName();
      final logo = await DbHelper.instance.logoBytes();
      final doc = pw.Document();
      for (final t in tbts) {
        final rows = atts[t['id'] as int] ?? [];
        doc.addPage(
          pw.MultiPage(
            pageFormat: PdfPageFormat.a4,
            build: (ctx) => [
              if (logo != null)
                pw.SizedBox(
                    height: 60,
                    child: pw.Image(pw.MemoryImage(logo),
                        fit: pw.BoxFit.contain)),
              pw.Text(company,
                  style: pw.TextStyle(
                      fontSize: 18, fontWeight: pw.FontWeight.bold)),
              pw.Text('Daily Toolbox Talk Record',
                  style: pw.TextStyle(
                      fontSize: 14, fontWeight: pw.FontWeight.bold)),
              pw.SizedBox(height: 8),
              pw.Text('Project: ${t['project']}'),
              pw.Text('Date: ${t['date']}'),
              pw.Text('Contractor: ${t['contractor']}'),
              pw.Text('Terminal: ${t['terminal']}'),
              pw.Text('Duty Shift: ${t['shift']}'),
              pw.Text('Location: ${t['location']}'),
              pw.SizedBox(height: 8),
              pw.Text('Topic: ${t['topic']}',
                  style: pw.TextStyle(fontWeight: pw.FontWeight.bold)),
              pw.Text('${t['topic_detail']}'),
              pw.SizedBox(height: 12),
              pw.Text('Attendees',
                  style: pw.TextStyle(fontWeight: pw.FontWeight.bold)),
              pw.Table.fromTextArray(
                headers: const [
                  'S.No',
                  'Iqama No.',
                  'Name',
                  'Designation',
                  'Nationality',
                  'Shift',
                  'Remarks'
                ],
                data: [
                  for (var i = 0; i < rows.length; i++)
                    [
                      '${i + 1}',
                      (rows[i]['iqama'] as String?) ?? '',
                      (rows[i]['name'] as String?) ?? '',
                      (rows[i]['designation'] as String?) ?? '',
                      (rows[i]['nationality'] as String?) ?? '',
                      (rows[i]['shift'] as String?) ?? '',
                      (rows[i]['remarks'] as String?) ?? '',
                    ],
                ],
              ),
              pw.SizedBox(height: 12),
              pw.Text('Signatures:',
                  style: pw.TextStyle(fontWeight: pw.FontWeight.bold)),
              pw.Wrap(
                spacing: 12,
                runSpacing: 8,
                children: [
                  for (final a in rows)
                    if (a['signature'] != null)
                      pw.Column(
                        children: [
                          pw.Image(
                              pw.MemoryImage(a['signature'] as Uint8List),
                              width: 90,
                              height: 42),
                          pw.Text((a['name'] as String?) ?? '',
                              style: const pw.TextStyle(fontSize: 9)),
                        ],
                      ),
                ],
              ),
            ],
          ),
        );
      }
      final bytes = await doc.save();
      final dir = await appDocs();
      final path = p.join(dir.path,
          'tbt_${fmtDate(_from)}_to_${fmtDate(_to)}.pdf');
      await File(path).writeAsBytes(bytes);
      if (mounted) {
        showDialog(
          context: context,
          builder: (ctx) => AlertDialog(
            title: const Text('PDF saved'),
            content: Text(path),
            actions: [
              TextButton(
                  onPressed: () => Navigator.pop(ctx),
                  child: const Text('OK')),
            ],
          ),
        );
      }
    } catch (e) {
      if (mounted) showMsg(context, 'PDF export failed: $e');
    } finally {
      if (mounted) setState(() => _busy = false);
    }
  }

  Future<void> _exportExcel() async {
    setState(() => _busy = true);
    try {
      final tbts = await _tbtsInRange();
      if (tbts.isEmpty) {
        showMsg(context, 'No TBT records in this range');
        return;
      }
      final ids = tbts.map((t) => t['id'] as int).toList();
      final atts = await _attendeesFor(ids);
      final company = await DbHelper.instance.companyName();
      final excel = Excel.createExcel();
      final sheet = excel['TBT'];
      sheet.appendRow([
        TextCellValue('Company'),
        TextCellValue('Date'),
        TextCellValue('Terminal'),
        TextCellValue('Shift'),
        TextCellValue('Project'),
        TextCellValue('Topic'),
        TextCellValue('S.No'),
        TextCellValue('Iqama No.'),
        TextCellValue('Name'),
        TextCellValue('Designation'),
        TextCellValue('Nationality'),
        TextCellValue('Attendee Shift'),
        TextCellValue('Remarks'),
      ]);
      for (final t in tbts) {
        final rows = atts[t['id'] as int] ?? [];
        if (rows.isEmpty) {
          sheet.appendRow([
            TextCellValue(company),
            TextCellValue('${t['date']}'),
            TextCellValue('${t['terminal']}'),
            TextCellValue('${t['shift']}'),
            TextCellValue('${t['project']}'),
            TextCellValue('${t['topic']}'),
            TextCellValue(''),
            TextCellValue(''),
            TextCellValue(''),
            TextCellValue(''),
            TextCellValue(''),
            TextCellValue(''),
            TextCellValue(''),
          ]);
        }
        for (var i = 0; i < rows.length; i++) {
          sheet.appendRow([
            TextCellValue(company),
            TextCellValue('${t['date']}'),
            TextCellValue('${t['terminal']}'),
            TextCellValue('${t['shift']}'),
            TextCellValue('${t['project']}'),
            TextCellValue('${t['topic']}'),
            TextCellValue('${i + 1}'),
            TextCellValue((rows[i]['iqama'] as String?) ?? ''),
            TextCellValue((rows[i]['name'] as String?) ?? ''),
            TextCellValue((rows[i]['designation'] as String?) ?? ''),
            TextCellValue((rows[i]['nationality'] as String?) ?? ''),
            TextCellValue((rows[i]['shift'] as String?) ?? ''),
            TextCellValue((rows[i]['remarks'] as String?) ?? ''),
          ]);
        }
      }
      final bytes = excel.save();
      if (bytes == null) throw 'Excel save returned null';
      final dir = await appDocs();
      final path = p.join(
          dir.path, 'tbt_${fmtDate(_from)}_to_${fmtDate(_to)}.xlsx');
      await File(path).writeAsBytes(bytes);
      if (mounted) {
        showDialog(
          context: context,
          builder: (ctx) => AlertDialog(
            title: const Text('Excel saved'),
            content: Text(path),
            actions: [
              TextButton(
                  onPressed: () => Navigator.pop(ctx),
                  child: const Text('OK')),
            ],
          ),
        );
      }
    } catch (e) {
      if (mounted) showMsg(context, 'Excel export failed: $e');
    } finally {
      if (mounted) setState(() => _busy = false);
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Export TBT')),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          ListTile(
            title: Text('From: ${fmtDate(_from)}'),
            trailing: const Icon(Icons.calendar_today),
            onTap: () async {
              final d = await showDatePicker(
                  context: context,
                  initialDate: _from,
                  firstDate: DateTime(2020),
                  lastDate: DateTime(2100));
              if (d != null) setState(() => _from = d);
            },
          ),
          ListTile(
            title: Text('To: ${fmtDate(_to)}'),
            trailing: const Icon(Icons.calendar_today),
            onTap: () async {
              final d = await showDatePicker(
                  context: context,
                  initialDate: _to,
                  firstDate: DateTime(2020),
                  lastDate: DateTime(2100));
              if (d != null) setState(() => _to = d);
            },
          ),
          Row(
            children: [
              TextButton(
                  onPressed: () {
                    setState(() {
                      _to = DateTime.now();
                      _from = _to;
                    });
                  },
                  child: const Text('1 day')),
              TextButton(
                  onPressed: () {
                    setState(() {
                      _to = DateTime.now();
                      _from = _to.subtract(const Duration(days: 6));
                    });
                  },
                  child: const Text('7 days')),
              TextButton(
                  onPressed: () {
                    setState(() {
                      _to = DateTime.now();
                      _from = _to.subtract(const Duration(days: 29));
                    });
                  },
                  child: const Text('1 month')),
            ],
          ),
          const SizedBox(height: 12),
          if (_busy)
            const Center(child: CircularProgressIndicator())
          else ...[
            ElevatedButton(
                onPressed: _exportPdf,
                child: const Text('Download PDF (official record)')),
            const SizedBox(height: 8),
            ElevatedButton(
                onPressed: _exportExcel,
                child: const Text('Download Excel (tabular data)')),
          ],
          const SizedBox(height: 16),
          const Text(
            'Files are saved inside the app storage; the full path is shown after export. Share the file from a file manager, or copy it from the app folder on the phone.',
            style: TextStyle(fontSize: 12, color: Colors.grey),
          ),
        ],
      ),
    );
  }
}

// PortSafe smoke test: home screen shows the four sections.
import 'package:flutter_test/flutter_test.dart';

import 'package:portsafe/main.dart';

void main() {
  testWidgets('Home shows the four sections', (WidgetTester tester) async {
    await tester.pumpWidget(const PortSafeApp());
    await tester.pumpAndSettle();
    expect(find.text('Toolbox Talk (TBT)'), findsOneWidget);
    expect(find.text('Attendance'), findsOneWidget);
    expect(find.text('Container Moves'), findsOneWidget);
    expect(find.text('Damage Report'), findsOneWidget);
  });
}

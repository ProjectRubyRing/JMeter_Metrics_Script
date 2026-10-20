"""検証用の WAR (perftest.war) を作る。

db.jsp: ExampleDS (H2) から接続を取り、指定ミリ秒保持してから返す。
        1 リクエストごとに数 MB を確保して GC を発生させ、HTTP セッションも作る。
使い方: python make_war.py <出力先ディレクトリ>
"""
import os
import sys
import zipfile

JSP = r'''<%@ page import="javax.naming.InitialContext, javax.sql.DataSource, java.sql.*" %><%
int ms = 50;
try { ms = Integer.parseInt(request.getParameter("ms")); } catch (Exception e) { }
int kb = 2048;
try { kb = Integer.parseInt(request.getParameter("kb")); } catch (Exception e) { }
byte[][] junk = new byte[kb / 64 + 1][];
for (int i = 0; i < junk.length; i++) { junk[i] = new byte[64 * 1024]; }
session.setAttribute("n", junk.length);
DataSource ds = (DataSource) new InitialContext().lookup("java:jboss/datasources/ExampleDS");
try (Connection c = ds.getConnection(); PreparedStatement ps = c.prepareStatement("SELECT 1")) {
    ResultSet rs = ps.executeQuery();
    rs.next();
    Thread.sleep(ms);
    out.println("ok " + rs.getInt(1) + " " + junk.length);
}
%>
'''

out_dir = sys.argv[1] if len(sys.argv) > 1 else '.'
path = os.path.join(out_dir, 'perftest.war')
with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('db.jsp', JSP)
    z.writestr('index.html', '<html><body>perftest</body></html>\n')
print(path)

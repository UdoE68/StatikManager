using System;
using System.IO;
using System.Text;

namespace StatikManager.Modules.Dokumente
{
    /// <summary>
    /// Erzeugt und verwaltet position.html für einen Positionsordner.
    /// – GeneriereHtmlMitXButtons: In-Memory-HTML für den WebBrowser (mit X-Buttons, NavigateToString-kompatibel)
    /// – LöscheAusschnitt: Löscht PNG+JSON, schreibt statische position.html neu
    /// </summary>
    internal static class PositionsHtmlHelper
    {
        // ── Öffentliche API ───────────────────────────────────────────────────

        /// <summary>
        /// Erzeugt HTML-String zur Anzeige im WebBrowser (NavigateToString):
        /// UTF-16-Charset, &lt;base href&gt; für relative Bildpfade, X-Button pro Ausschnitt.
        /// </summary>
        public static string GeneriereHtmlMitXButtons(string positionsPfad)
        {
            string[] pngDateien = HolePngDateien(positionsPfad);

            // Basis-URL für relative Bildpfade (daten/xxx.png)
            string baseUrl = "file:///"
                + positionsPfad.Replace('\\', '/').TrimEnd('/') + "/";

            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"de\">");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta http-equiv='Content-Type' content='text/html; charset=utf-16'>");
            sb.Append("<base href=\"").Append(HtmlAttr(baseUrl)).AppendLine("\">");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: Arial, sans-serif; max-width: 210mm; margin: 0 auto; padding: 10mm; }");
            sb.AppendLine("section { page-break-inside: avoid; margin-bottom: 20px; position: relative; }");
            sb.AppendLine("h3 { margin-bottom: 4px; padding-right: 36px; }");
            sb.AppendLine("img { max-width: 100%; border: 1px solid #ccc; display: block; }");
            sb.AppendLine(".massstab { font-size: 11px; color: #666; margin-top: 3px; }");
            sb.AppendLine("details { margin-top: 6px; font-size: 11px; color: #555; }");
            sb.AppendLine("summary { cursor: pointer; }");
            sb.AppendLine(".btn-x { position: absolute; top: 0; right: 0; width: 28px; height: 28px;");
            sb.AppendLine("  background: #cc3333; color: #fff; border: none; border-radius: 3px;");
            sb.AppendLine("  font-size: 18px; line-height: 28px; text-align: center; cursor: pointer;");
            sb.AppendLine("  font-family: Arial, sans-serif; }");
            sb.AppendLine(".btn-x:hover { background: #aa0000; }");
            sb.AppendLine("</style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");

            if (pngDateien.Length == 0)
            {
                sb.AppendLine("<p style='color:#888;font-style:italic'>Noch keine Ausschnitte vorhanden.</p>");
            }

            bool hatDatenOrdner = Directory.Exists(Path.Combine(positionsPfad, "daten"));

            foreach (string pngPfad in pngDateien)
            {
                string pngDateiname = Path.GetFileName(pngPfad);
                string basisname    = Path.GetFileNameWithoutExtension(pngPfad);
                string jsonPfad     = Path.Combine(Path.GetDirectoryName(pngPfad)!, basisname + ".json");
                string imgSrc       = hatDatenOrdner ? "daten/" + pngDateiname : pngDateiname;

                LesAusschnittJson(jsonPfad, basisname,
                    out string ueberschrift, out string massstab, out string datum, out string worldRect);

                string jsDateiname = EscapeJsString(pngDateiname);

                sb.AppendLine("<section>");
                sb.Append("  <button class=\"btn-x\" title=\"Ausschnitt l&#246;schen\"");
                sb.Append(" onclick=\"window.external.LoescheAusschnitt('").Append(jsDateiname).AppendLine("')\">&#xD7;</button>");
                sb.Append("  <h3>").Append(HtmlEncode(ueberschrift)).AppendLine("</h3>");
                sb.Append("  <img src=\"").Append(HtmlAttr(imgSrc)).Append("\" alt=\"").Append(HtmlAttr(ueberschrift)).AppendLine("\">");

                if (!string.IsNullOrEmpty(massstab))
                    sb.Append("  <div class=\"massstab\">Ma&#223;stab ").Append(HtmlEncode(massstab)).AppendLine("</div>");

                if (!string.IsNullOrEmpty(datum) || !string.IsNullOrEmpty(worldRect))
                {
                    sb.AppendLine("  <details>");
                    sb.AppendLine("    <summary>Metadaten</summary>");
                    if (!string.IsNullOrEmpty(datum))
                        sb.Append("    <div>Datum: ").Append(HtmlEncode(datum)).AppendLine("</div>");
                    if (!string.IsNullOrEmpty(worldRect))
                        sb.Append("    <div>Ausschnitt: ").Append(HtmlEncode(worldRect)).AppendLine("</div>");
                    sb.AppendLine("  </details>");
                }

                sb.AppendLine("</section>");
            }

            sb.AppendLine("</body>");
            sb.AppendLine("</html>");
            return sb.ToString();
        }

        /// <summary>
        /// Löscht PNG + zugehörige JSON aus daten/ und schreibt position.html (statische Version) neu.
        /// </summary>
        public static void LöscheAusschnitt(string positionsPfad, string pngDateiname)
        {
            string datenPfad   = Path.Combine(positionsPfad, "daten");
            bool hatDatenOrdner = Directory.Exists(datenPfad);

            string pngVollPfad = hatDatenOrdner
                ? Path.Combine(datenPfad, pngDateiname)
                : Path.Combine(positionsPfad, pngDateiname);

            if (File.Exists(pngVollPfad))
                File.Delete(pngVollPfad);

            string jsonDateiname = Path.ChangeExtension(pngDateiname, ".json");
            string jsonVollPfad  = hatDatenOrdner
                ? Path.Combine(datenPfad, jsonDateiname)
                : Path.Combine(positionsPfad, jsonDateiname);

            if (File.Exists(jsonVollPfad))
                File.Delete(jsonVollPfad);

            // Statische position.html auf Disk aktualisieren (ohne X-Buttons / ohne JS)
            SchreibeStatischeHtml(positionsPfad);
        }

        // ── Interne Hilfsmethoden ─────────────────────────────────────────────

        private static string[] HolePngDateien(string positionsPfad)
        {
            string datenPfad = Path.Combine(positionsPfad, "daten");
            string[] pngs = Directory.Exists(datenPfad)
                ? Directory.GetFiles(datenPfad, "*.png")
                : Directory.GetFiles(positionsPfad, "*.png");
            Array.Sort(pngs);
            return pngs;
        }

        /// <summary>
        /// Schreibt eine saubere, druckbare position.html auf Disk (UTF-8, kein JS) —
        /// identisch zur Ausgabe von PP_ZoomRahmen GeneriereHTML.
        /// </summary>
        private static void SchreibeStatischeHtml(string positionsPfad)
        {
            string[] pngDateien = HolePngDateien(positionsPfad);
            bool hatDatenOrdner = Directory.Exists(Path.Combine(positionsPfad, "daten"));

            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"de\">");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset=\"UTF-8\">");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: Arial, sans-serif; max-width: 210mm; margin: 0 auto; padding: 10mm; }");
            sb.AppendLine("section { page-break-inside: avoid; margin-bottom: 20px; }");
            sb.AppendLine("h3 { margin-bottom: 4px; }");
            sb.AppendLine("img { max-width: 100%; border: 1px solid #ccc; display: block; }");
            sb.AppendLine(".massstab { font-size: 11px; color: #666; margin-top: 3px; }");
            sb.AppendLine("details { margin-top: 6px; font-size: 11px; color: #555; }");
            sb.AppendLine("summary { cursor: pointer; }");
            sb.AppendLine("</style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");

            if (pngDateien.Length == 0)
                sb.AppendLine("<!-- Noch keine Ausschnitte -->");

            foreach (string pngPfad in pngDateien)
            {
                string pngDateiname = Path.GetFileName(pngPfad);
                string basisname    = Path.GetFileNameWithoutExtension(pngPfad);
                string jsonPfad     = Path.Combine(Path.GetDirectoryName(pngPfad)!, basisname + ".json");
                string imgSrc       = hatDatenOrdner ? "daten/" + pngDateiname : pngDateiname;

                LesAusschnittJson(jsonPfad, basisname,
                    out string ueberschrift, out string massstab, out string datum, out string worldRect);

                sb.AppendLine("<section>");
                sb.Append("  <h3>").Append(HtmlEncode(ueberschrift)).AppendLine("</h3>");
                sb.Append("  <img src=\"").Append(HtmlAttr(imgSrc)).Append("\" alt=\"").Append(HtmlAttr(ueberschrift)).AppendLine("\">");

                if (!string.IsNullOrEmpty(massstab))
                    sb.Append("  <div class=\"massstab\">Ma&#223;stab ").Append(HtmlEncode(massstab)).AppendLine("</div>");

                if (!string.IsNullOrEmpty(datum) || !string.IsNullOrEmpty(worldRect))
                {
                    sb.AppendLine("  <details>");
                    sb.AppendLine("    <summary>Metadaten</summary>");
                    if (!string.IsNullOrEmpty(datum))
                        sb.Append("    <div>Datum: ").Append(HtmlEncode(datum)).AppendLine("</div>");
                    if (!string.IsNullOrEmpty(worldRect))
                        sb.Append("    <div>Ausschnitt: ").Append(HtmlEncode(worldRect)).AppendLine("</div>");
                    sb.AppendLine("  </details>");
                }

                sb.AppendLine("</section>");
            }

            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            string htmlPfad = Path.Combine(positionsPfad, "position.html");
            File.WriteAllText(htmlPfad, sb.ToString(), Encoding.UTF8);
        }

        private static void LesAusschnittJson(string jsonPfad, string basisname,
            out string ueberschrift, out string massstab, out string datum, out string worldRect)
        {
            ueberschrift = basisname;
            massstab     = "";
            datum        = "";
            worldRect    = "";

            if (!File.Exists(jsonPfad)) return;

            string jText = File.ReadAllText(jsonPfad, Encoding.UTF8);

            string u = LeseJsonFeld(jText, "ueberschrift", "");
            if (!string.IsNullOrEmpty(u)) ueberschrift = u;

            string m = LeseJsonFeld(jText, "massstab", "");
            if (!string.IsNullOrEmpty(m)) massstab = m;

            string d = LeseJsonFeld(jText, "erstellt", "");
            if (!string.IsNullOrEmpty(d))
            {
                try   { datum = DateTime.Parse(d).ToString("dd.MM.yyyy HH:mm"); }
                catch { datum = d; }
            }

            string links  = LeseJsonFeld(jText, "links",  "");
            string unten  = LeseJsonFeld(jText, "unten",  "");
            string rechts = LeseJsonFeld(jText, "rechts", "");
            string oben   = LeseJsonFeld(jText, "oben",   "");
            if (!string.IsNullOrEmpty(links))
                worldRect = "L=" + links + " U=" + unten + " R=" + rechts + " O=" + oben + " m";
        }

        /// <summary>Minimaler JSON-Feld-Leser (kein externer Parser nötig).</summary>
        private static string LeseJsonFeld(string json, string key, string fallback)
        {
            if (string.IsNullOrEmpty(json)) return fallback;
            string suche = "\"" + key + "\"";
            int idx = json.IndexOf(suche, StringComparison.Ordinal);
            if (idx < 0) return fallback;
            int doppelpunkt = json.IndexOf(':', idx + suche.Length);
            if (doppelpunkt < 0) return fallback;
            int pos = doppelpunkt + 1;
            while (pos < json.Length && (json[pos] == ' ' || json[pos] == '\t' ||
                   json[pos] == '\r' || json[pos] == '\n'))
                pos++;
            if (pos >= json.Length) return fallback;
            if (json[pos] == '"')
            {
                int start = pos + 1, end = start;
                while (end < json.Length)
                {
                    if (json[end] == '"' && (end == 0 || json[end - 1] != '\\')) break;
                    end++;
                }
                if (end >= json.Length) return fallback;
                return json.Substring(start, end - start)
                    .Replace("\\\"", "\"").Replace("\\n", "\n").Replace("\\\\", "\\");
            }
            else
            {
                int start = pos, end = start;
                while (end < json.Length && json[end] != ',' && json[end] != '}' &&
                       json[end] != '\r' && json[end] != '\n')
                    end++;
                return json.Substring(start, end - start).Trim();
            }
        }

        private static string HtmlEncode(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length + 16);
            foreach (char c in s)
            {
                if (c == '<')       sb.Append("&lt;");
                else if (c == '>')  sb.Append("&gt;");
                else if (c == '&')  sb.Append("&amp;");
                else if (c > 127)   sb.Append("&#").Append((int)c).Append(';');
                else                sb.Append(c);
            }
            return sb.ToString();
        }

        private static string HtmlAttr(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length + 16);
            foreach (char c in s)
            {
                if (c == '"')       sb.Append("&quot;");
                else if (c == '<')  sb.Append("&lt;");
                else if (c == '>')  sb.Append("&gt;");
                else if (c == '&')  sb.Append("&amp;");
                else if (c > 127)   sb.Append("&#").Append((int)c).Append(';');
                else                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>Escaped einen String für ein JavaScript-String-Literal in einfachen Anführungszeichen.</summary>
        private static string EscapeJsString(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\")
                    .Replace("'",  "\\'")
                    .Replace("\r", "")
                    .Replace("\n", "");
        }
    }
}

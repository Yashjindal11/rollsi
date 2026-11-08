"""Generates manifest.xml with the ROLLSI ribbon tab. Run: python3 scripts/make-manifest.py"""
from pathlib import Path
from xml.sax.saxutils import escape

BASE = "https://localhost:3000"
ICONS = ["trend", "control", "waterfall", "target", "pareto", "gantt", "mekko", "diff", "cagr", "totals",
         "refresh", "remove", "style", "palette", "export", "demo", "panel"]

# (kind, id, label, tooltip, icon, action) ; action = function name, or ("pane",) ; menus carry items.
GROUPS = [
    ("Charts", "charts", [
        ("menu", "Trend", "Trend", "Line chart with a moving average. Select [dates] + values first.", "trend", [
            ("Trend7", "7-period moving average", "trend7"),
            ("Trend14", "14-period moving average", "trend14"),
            ("Trend30", "30-period moving average", "trend30"),
            ("Trend90", "90-period moving average", "trend90"),
        ]),
        ("menu", "Control", "Control", "Control chart with ±σ bands and outliers in red. Select [dates] + values.", "control", [
            ("Control2", "±2σ, whole series", "control2"),
            ("Control3", "±3σ, whole series", "control3"),
            ("Control3R", "±3σ, rolling 30 periods", "control3r30"),
        ]),
        ("button", "Waterfall", "Waterfall", "Bridge chart. Select labels + changes: first row = start, last = total, 'Subtotal' rows show running totals.", "waterfall", "waterfall"),
        ("button", "Target", "Actual vs Target", "Columns with target markers plus a variance chart. Select labels, actual, target.", "target", "actualVsTarget"),
        ("button", "Pareto", "Pareto", "Sorted bars with cumulative % and an 80% line. Select labels + values.", "pareto", "pareto"),
        ("button", "Gantt", "Gantt", "Timeline bars. Select task, start date, end date.", "gantt", "gantt"),
        ("button", "Mekko", "Mekko", "Marimekko chart. Select a table: categories across the top, segments down the side.", "mekko", "mekko"),
    ]),
    ("Annotate", "annotate", [
        ("menu", "Diff", "Difference", "Difference arrow on the selected column or line chart.", "diff", [
            ("DiffLast", "Last vs previous", "diffLast"),
            ("DiffFirstLast", "First vs last", "diffFirstLast"),
            ("DiffAll", "Every step", "diffAll"),
        ]),
        ("button", "Cagr", "CAGR", "CAGR arrow from the first to the last category of the selected chart.", "cagr", "cagr"),
        ("button", "Totals", "Totals", "Totals above each stack of the selected chart.", "totals", "totals"),
        ("button", "Refresh", "Refresh", "Redraw all annotations and Mekko charts on this sheet after data or size changes.", "refresh", "refresh"),
        ("button", "Remove", "Remove", "Remove annotations from the selected chart (or the whole sheet if no chart is selected).", "remove", "remove"),
    ]),
    ("Format", "format", [
        ("button", "Style", "Style", "Consultant style for the selected chart, or every chart on the sheet if none is selected.", "style", "style"),
        ("menu", "Palette", "Palette", "Colour palette for ROLLSI charts. Also restyles the selected chart.", "palette", [
            ("PalRollsi", "ROLLSI green", "paletteRollsi"),
            ("PalNavy", "Consulting navy", "paletteNavy"),
            ("PalMono", "Mono + accent", "paletteMono"),
            ("PalCb", "Colorblind safe", "paletteColorblind"),
            ("PalCustom", "Custom colours...", ("pane",)),
        ]),
    ]),
    ("Share", "share", [
        ("pane", "Export", "Export", "PNG images of charts at slide sizes, ready for PowerPoint.", "export", ("pane",)),
        ("button", "Demo", "Demo", "Create a demo sheet with sample data, functions and charts.", "demo", "demo"),
        ("pane", "Panel", "Settings & Help", "Palette, label settings and the =ROLLSI function reference.", "panel", ("pane",)),
    ]),
]

short_strings = {
    "Functions.Namespace": "ROLLSI",
    "GetStarted.Title": "ROLLSI is ready",
    "Tab.Label": "ROLLSI",
}
long_strings = {
    "GetStarted.Description": "Open the ROLLSI tab: select data and click a chart. Type =ROLLSI. in any cell for rolling-stat functions.",
}


def icon_xml(name, indent):
    pad = " " * indent
    return (f"{pad}<Icon>\n"
            + "".join(f'{pad}  <bt:Image size="{s}" resid="I.{name}.{s}" />\n' for s in (16, 32, 80))
            + f"{pad}</Icon>\n")


def action_xml(action, indent):
    pad = " " * indent
    if isinstance(action, tuple):
        return (f'{pad}<Action xsi:type="ShowTaskpane">\n{pad}  <TaskpaneId>RollsiPanel</TaskpaneId>\n'
                f'{pad}  <SourceLocation resid="Taskpane.Url" />\n{pad}</Action>\n')
    return f'{pad}<Action xsi:type="ExecuteFunction">\n{pad}  <FunctionName>{action}</FunctionName>\n{pad}</Action>\n'


def strings(key, label, tip):
    short_strings[f"L.{key}"] = label
    long_strings[f"T.{key}"] = tip


controls = []
for gname, gid, items in GROUPS:
    short_strings[f"G.{gid}"] = gname
    xml = f'            <Group id="Rollsi.{gid}">\n              <Label resid="G.{gid}" />\n'
    xml += icon_xml(items[0][4], 14)
    for kind, cid, label, tip, icon, action in items:
        strings(cid, label, tip)
        ctype = "Menu" if kind == "menu" else "Button"
        xml += f'              <Control xsi:type="{ctype}" id="Rollsi.{cid}">\n'
        xml += f'                <Label resid="L.{cid}" />\n'
        xml += f'                <Supertip>\n                  <Title resid="L.{cid}" />\n                  <Description resid="T.{cid}" />\n                </Supertip>\n'
        xml += icon_xml(icon, 16)
        if kind == "menu":
            xml += "                <Items>\n"
            for iid, ilabel, iaction in action:
                strings(iid, ilabel, f"{label}: {ilabel}")
                xml += f'                  <Item id="Rollsi.{iid}">\n'
                xml += f'                    <Label resid="L.{iid}" />\n'
                xml += f'                    <Supertip>\n                      <Title resid="L.{iid}" />\n                      <Description resid="T.{iid}" />\n                    </Supertip>\n'
                xml += icon_xml(icon, 20)
                xml += action_xml(iaction, 20)
                xml += "                  </Item>\n"
            xml += "                </Items>\n"
        else:
            xml += action_xml(action, 16)
        xml += "              </Control>\n"
    xml += "            </Group>\n"
    controls.append(xml)

images = "".join(
    f'        <bt:Image id="I.{n}.{s}" DefaultValue="{BASE}/assets/rb-{n}-{s}.png" />\n' for n in ICONS for s in (16, 32, 80)
)
images += "".join(f'        <bt:Image id="Icon.{s}" DefaultValue="{BASE}/assets/icon-{s}.png" />\n' for s in (16, 32, 80))
shorts = "".join(f'        <bt:String id="{k}" DefaultValue="{escape(v, {chr(34): "&quot;"})}" />\n' for k, v in short_strings.items())
longs = "".join(f'        <bt:String id="{k}" DefaultValue="{escape(v, {chr(34): "&quot;"})}" />\n' for k, v in long_strings.items())
assert all(len(v) <= 125 for v in short_strings.values()) and all(len(v) <= 250 for v in long_strings.values())

manifest = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!-- Generated by scripts/make-manifest.py. Edit that script, not this file. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>d6ec944c-9fbe-4e44-aa8b-edacf43f4045</Id>
  <Version>1.1.0.0</Version>
  <ProviderName>ROLLSI (open source)</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="ROLLSI" />
  <Description DefaultValue="think-cell-style charts, annotations and rolling-statistics functions for Excel." />
  <IconUrl DefaultValue="{BASE}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{BASE}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://github.com/Yashjindal11/rollsi"/>
  <AppDomains>
    <AppDomain>https://github.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="CustomFunctionsRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="{BASE}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="Functions.Script.Url" />
            </Script>
            <Page>
              <SourceLocation resid="Functions.Page.Url"/>
            </Page>
            <Metadata>
              <SourceLocation resid="Functions.Metadata.Url" />
            </Metadata>
            <Namespace resid="Functions.Namespace" />
          </ExtensionPoint>
        </AllFormFactors>
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="RollsiTab">
{"".join(controls)}              <Label resid="Tab.Label" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
{images}      </bt:Images>
      <bt:Urls>
        <bt:Url id="Functions.Script.Url" DefaultValue="{BASE}/public/functions.js" />
        <bt:Url id="Functions.Metadata.Url" DefaultValue="{BASE}/public/functions.json" />
        <bt:Url id="Functions.Page.Url" DefaultValue="{BASE}/public/functions.html" />
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://github.com/Yashjindal11/rollsi" />
        <bt:Url id="Commands.Url" DefaultValue="{BASE}/commands.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="{BASE}/taskpane.html" />
      </bt:Urls>
      <bt:ShortStrings>
{shorts}      </bt:ShortStrings>
      <bt:LongStrings>
{longs}      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
'''
Path(__file__).resolve().parent.parent.joinpath("manifest.xml").write_text(manifest)
print("manifest.xml written")

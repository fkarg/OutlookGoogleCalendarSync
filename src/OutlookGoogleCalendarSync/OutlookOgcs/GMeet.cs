using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using log4net;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    static class GMeet {
        private static readonly ILog log = LogManager.GetLogger(typeof(GMeet));

        private static String meetingIdToken = "GMEETURL";
        private static String plainInfo = "\r\nGoogle Meet joining information\r\nGMEETURL\r\nFirst time using Meet?  Learn more  <https://gsuite.google.com/learning-center/products/meet/get-started/>  \r\n\r\n";
        /// <summary>
        /// RTF document code for Google Meet details
        /// </summary>
        private static String rtfInfo =
        #region RTF document
            @"{\rtf1\ansi\ansicpg1252\deff0\deflang2057{\fonttbl{\f0\fnil\fcharset0 Calibri;}{\f1\fswiss\fprq2\fcharset0 Calibri;}}
{\colortbl ;\red0\green0\blue255;\red5\green99\blue193;}
{\*\generator Msftedit 5.41.21.2510;}\viewkind4\uc1\pard\lang9\f0\fs22{\pict\wmetafile8\picw2117\pich794\picwgoal1200\pichgoal450 
010009000003480e00000000320e00000000050000000b0200000000050000000c021a03450832
0e0000430f2000cc0000001e005000000000001a0345080000000028000000500000001e000000
0100180000000000201c0000c40e0000c40e00000000000000000000ffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
fffffffffffffffffffffffffffffffffffffffcf5efe79f60e38c40e38c40e38c40e38c4075c0
4075c04075c04075c04075c04075c04075c04075c04075c04075c04075c04080c650dcefcfffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffe79f60da6600da6600da6600da66
00da660047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac
0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe38c40da6600da66
00da6600da6600da660047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac
0047ac0047ac0047ac00ffffffffffffffffffaeda8fc5e4afffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe38c
40da6600da6600da6600da6600da660047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac
0047ac0047ac0047ac0047ac0047ac00fffffff3f9ef69bb3047ac0075c040ffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe2e1e0d9d7d6f5f5
f4ffffffffffffffffffffffffffffffffffffffffffffffffecebead9d7d6f5f5f4ffffffffff
ffffffffecebeabcbab8b3b0aeb3b0aecfceccffffffffffffffffffffffffffffffecebeabcba
b8b3b0aeb3b0aecfceccffffffffffffffffffffffffffffffecebeabcbab8bcbab8e2e1e0ffff
ffffffffe38c40da6600da6600da6600da6600da660047ac0047ac0047ac0047ac0047ac0047ac
0047ac0047ac0047ac0047ac0047ac0047ac0047ac00dcefcf52b11047ac0047ac0075c040ffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8d8a
8768635fd9d7d6ffffffffffffffffffffffffffffffffffffffffffffffffb3b0ae68635fd9d7
d6ffffffffffffc6c4c268635f68635f68635f68635f68635f8d8a87f5f5f4ffffffffffffc6c4
c2716c6968635f68635f68635f68635f84807df5f5f4ffffffffffffe2e1e0716c6968635f6863
5faaa7a5ffffffffffffe38c40da6600da6600da6600da6600da660047ac0047ac0047ac0047ac
0047ac0047ac0047ac0047ac0047ac0047ac0047ac0047ac003f9f0047ac0047ac0047ac0047ac
0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffff8d8a8768635fd9d7d6ffffffffffffffffffffffffffffffffffffffffffffffffb3b0
ae68635fd9d7d6ffffffd9d7d668635f68635fa09d9bd9d7d6c6c4c27a767368635f8d8a87ffff
ffe2e1e068635f68635f979391d9d7d6c6c4c284807d68635f84807dffffffffffff9793916863
5f84807db3b0aeecebeafffffffffffff69d55f47d1df47d1df47d1df47d1df47d1dd0d396d0ea
bfd0eabfd0eabfd0eabfd0eabfd0eabf69bb3047ac0047ac0047ac003792002d830047ac0047ac
0047ac0047ac0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffffffffffbcbab8cfceccffffffffff
ffffffffb3b0ae68635fd9d7d6ffffff97939168635fb3b0aefffffffffffffffffff5f5f48480
7dc6c4c2ffffff97939168635faaa7a5fffffffffffffffffff5f5f48d8a87c6c4c2ffffffffff
ff8d8a8768635fd9d7d6fffffffffffffffffffffffffca25cfc8426fc8426fc8426fc8426fc84
26fee0c8ffffffffffffffffffffffffffffffffffff75c04047ac0045a900328b002d83002d83
0047ac0047ac0047ac0047ac0075c040ffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffffcfcecc68635f716c
69f5f5f4ffffffffffffb3b0ae68635fd9d7d6ffffff68635f68635ff5f5f4ffffffffffffffff
ffffffffffffffffffffffffff716c6968635fecebeaffffffffffffffffffffffffffffffffff
ffffffffffffff8d8a8768635fd9d7d6fffffffffffffffffffffffffca25cfc8426fc8426fc84
26fc8426fc8426fee0c8ffffffffffffffffffffffffffffffffffff75c04041a2002f86002d83
002d83002d830047ac0047ac0047ac0047ac0075c040ffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffff7a76
7368635f68635faaa7a5ffffffffffffb3b0ae68635fd9d7d6ffffff68635f68635f68635f6863
5f68635f68635f68635f68635f68635fd9d7d668635f68635f68635f68635f68635f68635f6863
5f68635f68635fd9d7d6ffffff8d8a8768635fd9d7d6fffffffffffffffffffffffffca25cfc84
26fc8426fc8426fc8426fc8426fee0c8ffffffffffffffffffffffffffffffffffff6bb1402d83
002d83002d83002d83002d830047ac0047ac0047ac0047ac0075c040ffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffff8d8a8768635fd9d7d6ffff
ffb3b0ae68635f84807d716c6968635fecebeaffffffb3b0ae68635fd9d7d6ffffff68635f6863
5f84807d8d8a878d8a878d8a878d8a8768635f68635fffffff7a767368635f84807d8d8a878d8a
878d8a878d8a8768635f68635fe2e1e0ffffff8d8a8768635fd9d7d6ffffffffffffffffffffff
fffca25cfc8426fc8426fc8426fc8426fc8426fee0c8ffffffffffffffffffffffffffffffffff
ff51b79f2d83002d83002d83002d83002d830047ac0047ac0047ac0047ac0075c040ffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8d8a876863
5fd9d7d6f5f5f4716c6968635fe2e1e0b3b0ae68635f8d8a87ffffffb3b0ae68635fd9d7d6ffff
ff97939168635fa09d9bffffffffffffffffffc6c4c268635f7a7673ffffffa09d9b68635f9793
91ffffffffffffffffffd9d7d668635f716c69ffffffffffff8d8a8768635fd9d7d6ffffffffff
fffffffffffffffca25cfc8426fc8426fc8426fc8426fc8426fee0c8ffffffffffffffffffffff
ffffffffffffff40cbff08b0cf2a86102d83002d83002d830047ac0047ac0047ac0047ac0075c0
40ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ff8d8a8768635fd9d7d6a09d9b68635fa09d9bfffffff5f5f47a767368635fd9d7d6b3b0ae6863
5fd9d7d6ffffffe2e1e068635f68635f8d8a87b3b0aea09d9b68635f68635fcfceccffffffeceb
ea716c6968635f8d8a87b3b0aeaaa7a5716c6968635fbcbab8e2e1e08d8a87716c6968635f8480
7d8d8a87c6c4c2fffffffffffffca25cfc8426fc8426fc8426fc8426fc8426fee0c8ffffffffff
ffffffffffffffffffffffffff40cbff00baff08b0cf258d302d83002d830047ac0047ac0047ac
0047ac0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffff8d8a8768635fbcbab868635f716c69f5f5f4ffffffffffffc6c4c268635f7a76
73b3b0ae68635fd9d7d6ffffffffffffcfcecc716c6968635f68635f68635f68635fb3b0aeffff
ffffffffffffffe2e1e07a767368635f68635f68635f68635faaa7a5ffffffd9d7d668635f6863
5f68635f68635f68635fb3b0aeffffffffffffd79681ca7457ca7457ca7457ca7457ca74577fcd
e47fdcff7fdcff7fdcff7fdcff7fdcff7fdcff20c2ff00baff00baff03b7ef2291402d830047ac
0047ac0047ac0047ac0075c040ffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffff8d8a8768635f716c6968635fb3b0aeffffffffffffffffffffff
ff8d8a8768635f7a767368635fd9d7d6fffffffffffffffffff5f5f4d9d7d6b3b0aecfcecceceb
eaffffffffffffffffffffffffffffffffffffd9d7d6b3b0aecfceccecebeaffffffffffffffff
ffffffff8d8a8768635fd9d7d6ffffffffffffffffffffffffd8dbfb414eeb3543ea3543ea3543
ea3543ea00baff00baff00baff00baff00baff00baff00baff00baff00baff00baff00baff00ba
ff199b7047ac0047ac0047ac0047ac0075c040ffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffff8d8a8768635f68635f7a7673ffffffffffffffff
ffffffffffffffd9d7d668635f68635f68635fd9d7d6ffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffffffffffffffffffffffd8dbfb414e
eb3543ea3543ea3543ea00baff00baff00baff00baff00baff00baff00baff00baff00baff00ba
ff00baff00baff00baffaeda8f47ac0047ac0047ac0075c040ffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffff8d8a8768635f68635fcfceccffff
ffffffffffffffffffffffffffffffffa09d9b68635f68635fd9d7d6ffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffff8d8a8768635fd9d7d6ffffffffffffffffffffffffffff
ffffffffd8dbfb414eeb3543ea3543ea00baff00baff00baff00baff00baff00baff00baff00ba
ff00baff00baff00baff00baff00baffffffffdcefcf52b11047ac0075c040ffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8d8a8768635f8d8a
87ffffffffffffffffffffffffffffffffffffffffffecebea716c6968635fd9d7d6ffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffd8dbfb414eeb3543ea00baff00baff00baff00baff00baff00ba
ff00baff00baff00baff00baff00baff00baff00bafffffffffffffff3f9ef80c650a2d57fffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe2e1
e0d9d7d6ecebeaffffffffffffffffffffffffffffffffffffffffffffffffe2e1e0d9d7d6f5f5
f4ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffd8dbfb414eeb00baff00baff00baff00ba
ff00baff00baff00baff00baff00baff00baff00baff00baff40cbffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffd8dbfb40cbff40cb
ff40cbff40cbff40cbff40cbff40cbff40cbff40cbff40cbff40cbff60d3ffeffaffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
ffffffffffffffffffffffffff030000000000
}\lang2057\f1\par
\b Google Meet joining information\b0\par
{\field{\*\fldinst{HYPERLINK ""GMEETURL""}}{\fldrslt{\ul\cf1 GMEETURL}}}\f1\fs22\par
First time using Meet?{\field{\*\fldinst{ HYPERLINK ""https://gsuite.google.com/learning-center/products/meet/get-started/"" \\\\l ""!/section-2-3?hl=en-GB"" \\\\t ""_blank"" } } {\fldrslt{\cf2\ul\b\~Learn more} } }\cf0\ulnone\b0\f1\fs22\par
\par
\pard\sa200\sl276\slmult1\lang9\f0\par
}";
        #endregion

        public static String PlainInfo(String meetingUrl) {
            return plainInfo.Replace(meetingIdToken, meetingUrl);
        }

        public static String RtfInfo(String meetingUrl) {
            return rtfInfo.Replace(meetingIdToken, meetingUrl);
        }

        public static String GetInfoBlock(AppointmentItem ai) {
            try {
                OlBodyFormat format = (OlBodyFormat)ai.GetType().InvokeMember("BodyFormat", System.Reflection.BindingFlags.GetProperty, null, ai, null);
                if (format == OlBodyFormat.olFormatUnspecified) {
                    log.Warn("Unspecified body format!");
                    return "";

                } else if (format == OlBodyFormat.olFormatHTML) {
                    log.Warn("The body format is HTML; unsupported for GMeet code sync.");
                    return "";

                } else if (format == OlBodyFormat.olFormatPlain) {
                    return ai.Body;

                } else if (format == OlBodyFormat.olFormatRichText) {
                    String bodyCode = Encoding.ASCII.GetString(ai.RTFBody as byte[]);
                    return bodyCode;
                }

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }

            //if (bodyIsHtml(ai.RTFBody)) 
            return null;
        }

        /// <summary>
        /// Add/update Google Meet information block to Outlook appointment body.
        /// </summary>
        /// <param name="ai">The appointment to update</param>
        /// <param name="gMeetUrl">The URL of the Meeting</param>
        public static void GoogleMeet(this Microsoft.Office.Interop.Outlook.AppointmentItem ai, String gMeetUrl) {
            if (String.IsNullOrEmpty(gMeetUrl)) {
                CustomProperty.Remove(ref ai, CustomProperty.MetadataId.gMeetUrl);

            } else {
                CustomProperty.Add(ref ai, CustomProperty.MetadataId.gMeetUrl, gMeetUrl);
                Regex rgx = new Regex(@"https:\/\/meet\.google\.com\/[a-z]{3}-[a-z]{4}-[a-z]{3}", RegexOptions.None);

                if (ai.BodyFormat() == OlBodyFormat.olFormatPlain) {
                    if (String.IsNullOrEmpty(ai.Body?.Replace(PlainInfo(gMeetUrl + "ZZZ"), ""))) {
                        log.Debug("Adding GMeet RTF body to Outlook");
                        Calendar.Instance.IOutlook.AddRtfBody(ref ai, RtfInfo(gMeetUrl));
                    } else {
                        log.Debug("Updating GMeet plaintext body in Outlook");
                        ai.Body = rgx.Replace(ai.Body, gMeetUrl);
                    }
                } else if (ai.BodyFormat() == OlBodyFormat.olFormatRichText) {
                    if (String.IsNullOrEmpty(ai.Body)) {
                        log.Debug("Adding GMeet RTF body to Outlook");
                        Calendar.Instance.IOutlook.AddRtfBody(ref ai, RtfInfo(gMeetUrl));
                    } else {
                        log.Debug("Updating GMeet RTF body in Outlook");
                        String newRtfBody = rgx.Replace(ai.RTFBodyAsString(), gMeetUrl);
                        Calendar.Instance.IOutlook.AddRtfBody(ref ai, newRtfBody);
                    }
                } else {
                    if (String.IsNullOrEmpty(ai.Body)) {
                        log.Debug("Adding GMeet RTF body to Outlook");
                        Calendar.Instance.IOutlook.AddRtfBody(ref ai, RtfInfo(gMeetUrl));
                    } else {
                        log.Warn(ai.BodyFormat().ToString() + " is not fully supported. Attempting update of pre-existing GMeet URL.");
                        String newHtmlBody = rgx.Replace(ai.RTFBodyAsString(), gMeetUrl);
                        Calendar.Instance.IOutlook.AddRtfBody(ref ai, newHtmlBody);
                    }
                }
            }
        }
    }
}

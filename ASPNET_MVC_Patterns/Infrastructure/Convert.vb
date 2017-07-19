﻿Imports System.IO

Public Class Convert

    ''' <summary>
    ''' Get content type.
    ''' </summary>
    ''' <param name="filename">String</param>
    ''' <returns>
    ''' Reference: https://en.wikipedia.org/wiki/Internet_media_type</returns>
    Public Shared Function GetContentType(filename As String) As String
        Dim extension As String = Path.GetExtension(filename)
        If extension Is Nothing Then
            Throw New ArgumentNullException("extension")
        End If
        If extension.StartsWith(".") Then
            extension = extension.Substring(1)
        End If
        ' List from here: http://stackoverflow.com/questions/1029740/get-mime-type-from-filename-extension/14108040#14108040
        Select Case extension.ToLower()
            Case "323"
                Return "text/h323"
            Case "3g2"
                Return "video/3gpp2"
            Case "3gp"
                Return "video/3gpp"
            Case "3gp2"
                Return "video/3gpp2"
            Case "3gpp"
                Return "video/3gpp"
            Case "7z"
                Return "application/x-7z-compressed"
            Case "aa"
                Return "audio/audible"
            Case "aac"
                Return "audio/aac"
            Case "aaf"
                Return "application/octet-stream"
            Case "aax"
                Return "audio/vnd.audible.aax"
            Case "ac3"
                Return "audio/ac3"
            Case "aca"
                Return "application/octet-stream"
            Case "accda"
                Return "application/msaccess.addin"
            Case "accdb"
                Return "application/msaccess"
            Case "accdc"
                Return "application/msaccess.cab"
            Case "accde"
                Return "application/msaccess"
            Case "accdr"
                Return "application/msaccess.runtime"
            Case "accdt"
                Return "application/msaccess"
            Case "accdw"
                Return "application/msaccess.webapplication"
            Case "accft"
                Return "application/msaccess.ftemplate"
            Case "acx"
                Return "application/internet-property-stream"
            Case "addin"
                Return "text/xml"
            Case "ade"
                Return "application/msaccess"
            Case "adobebridge"
                Return "application/x-bridge-url"
            Case "adp"
                Return "application/msaccess"
            Case "adt"
                Return "audio/vnd.dlna.adts"
            Case "adts"
                Return "audio/aac"
            Case "afm"
                Return "application/octet-stream"
            Case "ai"
                Return "application/postscript"
            Case "aif"
                Return "audio/x-aiff"
            Case "aifc"
                Return "audio/aiff"
            Case "aiff"
                Return "audio/aiff"
            Case "air"
                Return "application/vnd.adobe.air-application-installer-package+zip"
            Case "amc"
                Return "application/x-mpeg"
            Case "application"
                Return "application/x-ms-application"
            Case "art"
                Return "image/x-jg"
            Case "asa"
                Return "application/xml"
            Case "asax"
                Return "application/xml"
            Case "ascx"
                Return "application/xml"
            Case "asd"
                Return "application/octet-stream"
            Case "asf"
                Return "video/x-ms-asf"
            Case "ashx"
                Return "application/xml"
            Case "asi"
                Return "application/octet-stream"
            Case "asm"
                Return "text/plain"
            Case "asmx"
                Return "application/xml"
            Case "aspx"
                Return "application/xml"
            Case "asr"
                Return "video/x-ms-asf"
            Case "asx"
                Return "video/x-ms-asf"
            Case "atom"
                Return "application/atom+xml"
            Case "au"
                Return "audio/basic"
            Case "avi"
                Return "video/x-msvideo"
            Case "axs"
                Return "application/olescript"
            Case "bas"
                Return "text/plain"
            Case "bcpio"
                Return "application/x-bcpio"
            Case "bin"
                Return "application/octet-stream"
            Case "bmp"
                Return "image/bmp"
            Case "c"
                Return "text/plain"
            Case "cab"
                Return "application/octet-stream"
            Case "caf"
                Return "audio/x-caf"
            Case "calx"
                Return "application/vnd.ms-office.calx"
            Case "cat"
                Return "application/vnd.ms-pki.seccat"
            Case "cc"
                Return "text/plain"
            Case "cd"
                Return "text/plain"
            Case "cdda"
                Return "audio/aiff"
            Case "cdf"
                Return "application/x-cdf"
            Case "cer"
                Return "application/x-x509-ca-cert"
            Case "chm"
                Return "application/octet-stream"
            Case "class"
                Return "application/x-java-applet"
            Case "clp"
                Return "application/x-msclip"
            Case "cmx"
                Return "image/x-cmx"
            Case "cnf"
                Return "text/plain"
            Case "cod"
                Return "image/cis-cod"
            Case "config"
                Return "application/xml"
            Case "contact"
                Return "text/x-ms-contact"
            Case "coverage"
                Return "application/xml"
            Case "cpio"
                Return "application/x-cpio"
            Case "cpp"
                Return "text/plain"
            Case "crd"
                Return "application/x-mscardfile"
            Case "crl"
                Return "application/pkix-crl"
            Case "crt"
                Return "application/x-x509-ca-cert"
            Case "cs"
                Return "text/plain"
            Case "csdproj"
                Return "text/plain"
            Case "csh"
                Return "application/x-csh"
            Case "csproj"
                Return "text/plain"
            Case "css"
                Return "text/css"
            Case "csv"
                Return "text/csv"
            Case "cur"
                Return "application/octet-stream"
            Case "cxx"
                Return "text/plain"
            Case "dat"
                Return "application/octet-stream"
            Case "datasource"
                Return "application/xml"
            Case "dbproj"
                Return "text/plain"
            Case "dcr"
                Return "application/x-director"
            Case "def"
                Return "text/plain"
            Case "deploy"
                Return "application/octet-stream"
            Case "der"
                Return "application/x-x509-ca-cert"
            Case "dgml"
                Return "application/xml"
            Case "dib"
                Return "image/bmp"
            Case "dif"
                Return "video/x-dv"
            Case "dir"
                Return "application/x-director"
            Case "disco"
                Return "text/xml"
            Case "dll"
                Return "application/x-msdownload"
            Case "dll.config"
                Return "text/xml"
            Case "dlm"
                Return "text/dlm"
            Case "doc"
                Return "application/msword"
            Case "docm"
                Return "application/vnd.ms-word.document.macroenabled.12"
            Case "docx"
                Return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            Case "dot"
                Return "application/msword"
            Case "dotm"
                Return "application/vnd.ms-word.template.macroenabled.12"
            Case "dotx"
                Return "application/vnd.openxmlformats-officedocument.wordprocessingml.template"
            Case "dsp"
                Return "application/octet-stream"
            Case "dsw"
                Return "text/plain"
            Case "dtd"
                Return "text/xml"
            Case "dtsconfig"
                Return "text/xml"
            Case "dv"
                Return "video/x-dv"
            Case "dvi"
                Return "application/x-dvi"
            Case "dwf"
                Return "drawing/x-dwf"
            Case "dwp"
                Return "application/octet-stream"
            Case "dxr"
                Return "application/x-director"
            Case "eml"
                Return "message/rfc822"
            Case "emz"
                Return "application/octet-stream"
            Case "eot"
                Return "application/octet-stream"
            Case "eps"
                Return "application/postscript"
            Case "etl"
                Return "application/etl"
            Case "etx"
                Return "text/x-setext"
            Case "evy"
                Return "application/envoy"
            Case "exe"
                Return "application/octet-stream"
            Case "exe.config"
                Return "text/xml"
            Case "fdf"
                Return "application/vnd.fdf"
            Case "fif"
                Return "application/fractals"
            Case "filters"
                Return "application/xml"
            Case "fla"
                Return "application/octet-stream"
            Case "flr"
                Return "x-world/x-vrml"
            Case "flv"
                Return "video/x-flv"
            Case "fsscript"
                Return "application/fsharp-script"
            Case "fsx"
                Return "application/fsharp-script"
            Case "generictest"
                Return "application/xml"
            Case "gif"
                Return "image/gif"
            Case "group"
                Return "text/x-ms-group"
            Case "gsm"
                Return "audio/x-gsm"
            Case "gtar"
                Return "application/x-gtar"
            Case "gz"
                Return "application/x-gzip"
            Case "h"
                Return "text/plain"
            Case "hdf"
                Return "application/x-hdf"
            Case "hdml"
                Return "text/x-hdml"
            Case "hhc"
                Return "application/x-oleobject"
            Case "hhk"
                Return "application/octet-stream"
            Case "hhp"
                Return "application/octet-stream"
            Case "hlp"
                Return "application/winhlp"
            Case "hpp"
                Return "text/plain"
            Case "hqx"
                Return "application/mac-binhex40"
            Case "hta"
                Return "application/hta"
            Case "htc"
                Return "text/x-component"
            Case "htm"
                Return "text/html"
            Case "html"
                Return "text/html"
            Case "htt"
                Return "text/webviewhtml"
            Case "hxa"
                Return "application/xml"
            Case "hxc"
                Return "application/xml"
            Case "hxd"
                Return "application/octet-stream"
            Case "hxe"
                Return "application/xml"
            Case "hxf"
                Return "application/xml"
            Case "hxh"
                Return "application/octet-stream"
            Case "hxi"
                Return "application/octet-stream"
            Case "hxk"
                Return "application/xml"
            Case "hxq"
                Return "application/octet-stream"
            Case "hxr"
                Return "application/octet-stream"
            Case "hxs"
                Return "application/octet-stream"
            Case "hxt"
                Return "text/html"
            Case "hxv"
                Return "application/xml"
            Case "hxw"
                Return "application/octet-stream"
            Case "hxx"
                Return "text/plain"
            Case "i"
                Return "text/plain"
            Case "ico"
                Return "image/x-icon"
            Case "ics"
                Return "application/octet-stream"
            Case "idl"
                Return "text/plain"
            Case "ief"
                Return "image/ief"
            Case "iii"
                Return "application/x-iphone"
            Case "inc"
                Return "text/plain"
            Case "inf"
                Return "application/octet-stream"
            Case "inl"
                Return "text/plain"
            Case "ins"
                Return "application/x-internet-signup"
            Case "ipa"
                Return "application/x-itunes-ipa"
            Case "ipg"
                Return "application/x-itunes-ipg"
            Case "ipproj"
                Return "text/plain"
            Case "ipsw"
                Return "application/x-itunes-ipsw"
            Case "iqy"
                Return "text/x-ms-iqy"
            Case "isp"
                Return "application/x-internet-signup"
            Case "ite"
                Return "application/x-itunes-ite"
            Case "itlp"
                Return "application/x-itunes-itlp"
            Case "itms"
                Return "application/x-itunes-itms"
            Case "itpc"
                Return "application/x-itunes-itpc"
            Case "ivf"
                Return "video/x-ivf"
            Case "jar"
                Return "application/java-archive"
            Case "java"
                Return "application/octet-stream"
            Case "jck"
                Return "application/liquidmotion"
            Case "jcz"
                Return "application/liquidmotion"
            Case "jfif"
                Return "image/pjpeg"
            Case "jnlp"
                Return "application/x-java-jnlp-file"
            Case "jpb"
                Return "application/octet-stream"
            Case "jpe"
                Return "image/jpeg"
            Case "jpeg"
                Return "image/jpeg"
            Case "jpg"
                Return "image/jpeg"
            Case "js"
                Return "application/x-javascript"
            Case "jsx"
                Return "text/jscript"
            Case "jsxbin"
                Return "text/plain"
            Case "latex"
                Return "application/x-latex"
            Case "library-ms"
                Return "application/windows-library+xml"
            Case "lit"
                Return "application/x-ms-reader"
            Case "loadtest"
                Return "application/xml"
            Case "lpk"
                Return "application/octet-stream"
            Case "lsf"
                Return "video/x-la-asf"
            Case "lst"
                Return "text/plain"
            Case "lsx"
                Return "video/x-la-asf"
            Case "lzh"
                Return "application/octet-stream"
            Case "m13"
                Return "application/x-msmediaview"
            Case "m14"
                Return "application/x-msmediaview"
            Case "m1v"
                Return "video/mpeg"
            Case "m2t"
                Return "video/vnd.dlna.mpeg-tts"
            Case "m2ts"
                Return "video/vnd.dlna.mpeg-tts"
            Case "m2v"
                Return "video/mpeg"
            Case "m3u"
                Return "audio/x-mpegurl"
            Case "m3u8"
                Return "audio/x-mpegurl"
            Case "m4a"
                Return "audio/m4a"
            Case "m4b"
                Return "audio/m4b"
            Case "m4p"
                Return "audio/m4p"
            Case "m4r"
                Return "audio/x-m4r"
            Case "m4v"
                Return "video/x-m4v"
            Case "mac"
                Return "image/x-macpaint"
            Case "mak"
                Return "text/plain"
            Case "man"
                Return "application/x-troff-man"
            Case "manifest"
                Return "application/x-ms-manifest"
            Case "map"
                Return "text/plain"
            Case "master"
                Return "application/xml"
            Case "mda"
                Return "application/msaccess"
            Case "mdb"
                Return "application/x-msaccess"
            Case "mde"
                Return "application/msaccess"
            Case "mdp"
                Return "application/octet-stream"
            Case "me"
                Return "application/x-troff-me"
            Case "mfp"
                Return "application/x-shockwave-flash"
            Case "mht"
                Return "message/rfc822"
            Case "mhtml"
                Return "message/rfc822"
            Case "mid"
                Return "audio/mid"
            Case "midi"
                Return "audio/mid"
            Case "mix"
                Return "application/octet-stream"
            Case "mk"
                Return "text/plain"
            Case "mmf"
                Return "application/x-smaf"
            Case "mno"
                Return "text/xml"
            Case "mny"
                Return "application/x-msmoney"
            Case "mod"
                Return "video/mpeg"
            Case "mov"
                Return "video/quicktime"
            Case "movie"
                Return "video/x-sgi-movie"
            Case "mp2"
                Return "video/mpeg"
            Case "mp2v"
                Return "video/mpeg"
            Case "mp3"
                Return "audio/mpeg"
            Case "mp4"
                Return "video/mp4"
            Case "mp4v"
                Return "video/mp4"
            Case "mpa"
                Return "video/mpeg"
            Case "mpe"
                Return "video/mpeg"
            Case "mpeg"
                Return "video/mpeg"
            Case "mpf"
                Return "application/vnd.ms-mediapackage"
            Case "mpg"
                Return "video/mpeg"
            Case "mpp"
                Return "application/vnd.ms-project"
            Case "mpv2"
                Return "video/mpeg"
            Case "mqv"
                Return "video/quicktime"
            Case "ms"
                Return "application/x-troff-ms"
            Case "msi"
                Return "application/octet-stream"
            Case "mso"
                Return "application/octet-stream"
            Case "mts"
                Return "video/vnd.dlna.mpeg-tts"
            Case "mtx"
                Return "application/xml"
            Case "mvb"
                Return "application/x-msmediaview"
            Case "mvc"
                Return "application/x-miva-compiled"
            Case "mxp"
                Return "application/x-mmxp"
            Case "nc"
                Return "application/x-netcdf"
            Case "nsc"
                Return "video/x-ms-asf"
            Case "nws"
                Return "message/rfc822"
            Case "ocx"
                Return "application/octet-stream"
            Case "oda"
                Return "application/oda"
            Case "odc"
                Return "text/x-ms-odc"
            Case "odh"
                Return "text/plain"
            Case "odl"
                Return "text/plain"
            Case "odp"
                Return "application/vnd.oasis.opendocument.presentation"
            Case "ods"
                Return "application/oleobject"
            Case "odt"
                Return "application/vnd.oasis.opendocument.text"
            Case "one"
                Return "application/onenote"
            Case "onea"
                Return "application/onenote"
            Case "onepkg"
                Return "application/onenote"
            Case "onetmp"
                Return "application/onenote"
            Case "onetoc"
                Return "application/onenote"
            Case "onetoc2"
                Return "application/onenote"
            Case "orderedtest"
                Return "application/xml"
            Case "osdx"
                Return "application/opensearchdescription+xml"
            Case "p10"
                Return "application/pkcs10"
            Case "p12"
                Return "application/x-pkcs12"
            Case "p7b"
                Return "application/x-pkcs7-certificates"
            Case "p7c"
                Return "application/pkcs7-mime"
            Case "p7m"
                Return "application/pkcs7-mime"
            Case "p7r"
                Return "application/x-pkcs7-certreqresp"
            Case "p7s"
                Return "application/pkcs7-signature"
            Case "pbm"
                Return "image/x-portable-bitmap"
            Case "pcast"
                Return "application/x-podcast"
            Case "pct"
                Return "image/pict"
            Case "pcx"
                Return "application/octet-stream"
            Case "pcz"
                Return "application/octet-stream"
            Case "pdf"
                Return "application/pdf"
            Case "pfb"
                Return "application/octet-stream"
            Case "pfm"
                Return "application/octet-stream"
            Case "pfx"
                Return "application/x-pkcs12"
            Case "pgm"
                Return "image/x-portable-graymap"
            Case "pic"
                Return "image/pict"
            Case "pict"
                Return "image/pict"
            Case "pkgdef"
                Return "text/plain"
            Case "pkgundef"
                Return "text/plain"
            Case "pko"
                Return "application/vnd.ms-pki.pko"
            Case "pls"
                Return "audio/scpls"
            Case "pma"
                Return "application/x-perfmon"
            Case "pmc"
                Return "application/x-perfmon"
            Case "pml"
                Return "application/x-perfmon"
            Case "pmr"
                Return "application/x-perfmon"
            Case "pmw"
                Return "application/x-perfmon"
            Case "png"
                Return "image/png"
            Case "pnm"
                Return "image/x-portable-anymap"
            Case "pnt"
                Return "image/x-macpaint"
            Case "pntg"
                Return "image/x-macpaint"
            Case "pnz"
                Return "image/png"
            Case "pot"
                Return "application/vnd.ms-powerpoint"
            Case "potm"
                Return "application/vnd.ms-powerpoint.template.macroenabled.12"
            Case "potx"
                Return "application/vnd.openxmlformats-officedocument.presentationml.template"
            Case "ppa"
                Return "application/vnd.ms-powerpoint"
            Case "ppam"
                Return "application/vnd.ms-powerpoint.addin.macroenabled.12"
            Case "ppm"
                Return "image/x-portable-pixmap"
            Case "pps"
                Return "application/vnd.ms-powerpoint"
            Case "ppsm"
                Return "application/vnd.ms-powerpoint.slideshow.macroenabled.12"
            Case "ppsx"
                Return "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
            Case "ppt"
                Return "application/vnd.ms-powerpoint"
            Case "pptm"
                Return "application/vnd.ms-powerpoint.presentation.macroenabled.12"
            Case "pptx"
                Return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            Case "prf"
                Return "application/pics-rules"
            Case "prm"
                Return "application/octet-stream"
            Case "prx"
                Return "application/octet-stream"
            Case "ps"
                Return "application/postscript"
            Case "psc1"
                Return "application/powershell"
            Case "psd"
                Return "application/octet-stream"
            Case "psess"
                Return "application/xml"
            Case "psm"
                Return "application/octet-stream"
            Case "psp"
                Return "application/octet-stream"
            Case "pub"
                Return "application/x-mspublisher"
            Case "pwz"
                Return "application/vnd.ms-powerpoint"
            Case "qht"
                Return "text/x-html-insertion"
            Case "qhtm"
                Return "text/x-html-insertion"
            Case "qt"
                Return "video/quicktime"
            Case "qti"
                Return "image/x-quicktime"
            Case "qtif"
                Return "image/x-quicktime"
            Case "qtl"
                Return "application/x-quicktimeplayer"
            Case "qxd"
                Return "application/octet-stream"
            Case "ra"
                Return "audio/x-pn-realaudio"
            Case "ram"
                Return "audio/x-pn-realaudio"
            Case "rar"
                Return "application/octet-stream"
            Case "ras"
                Return "image/x-cmu-raster"
            Case "rat"
                Return "application/rat-file"
            Case "rc"
                Return "text/plain"
            Case "rc2"
                Return "text/plain"
            Case "rct"
                Return "text/plain"
            Case "rdlc"
                Return "application/xml"
            Case "resx"
                Return "application/xml"
            Case "rf"
                Return "image/vnd.rn-realflash"
            Case "rgb"
                Return "image/x-rgb"
            Case "rgs"
                Return "text/plain"
            Case "rm"
                Return "application/vnd.rn-realmedia"
            Case "rmi"
                Return "audio/mid"
            Case "rmp"
                Return "application/vnd.rn-rn_music_package"
            Case "roff"
                Return "application/x-troff"
            Case "rpm"
                Return "audio/x-pn-realaudio-plugin"
            Case "rqy"
                Return "text/x-ms-rqy"
            Case "rtf"
                Return "application/rtf"
            Case "rtx"
                Return "text/richtext"
            Case "ruleset"
                Return "application/xml"
            Case "s"
                Return "text/plain"
            Case "safariextz"
                Return "application/x-safari-safariextz"
            Case "scd"
                Return "application/x-msschedule"
            Case "sct"
                Return "text/scriptlet"
            Case "sd2"
                Return "audio/x-sd2"
            Case "sdp"
                Return "application/sdp"
            Case "sea"
                Return "application/octet-stream"
            Case "searchconnector-ms"
                Return "application/windows-search-connector+xml"
            Case "setpay"
                Return "application/set-payment-initiation"
            Case "setreg"
                Return "application/set-registration-initiation"
            Case "settings"
                Return "application/xml"
            Case "sgimb"
                Return "application/x-sgimb"
            Case "sgml"
                Return "text/sgml"
            Case "sh"
                Return "application/x-sh"
            Case "shar"
                Return "application/x-shar"
            Case "shtml"
                Return "text/html"
            Case "sit"
                Return "application/x-stuffit"
            Case "sitemap"
                Return "application/xml"
            Case "skin"
                Return "application/xml"
            Case "sldm"
                Return "application/vnd.ms-powerpoint.slide.macroenabled.12"
            Case "sldx"
                Return "application/vnd.openxmlformats-officedocument.presentationml.slide"
            Case "slk"
                Return "application/vnd.ms-excel"
            Case "sln"
                Return "text/plain"
            Case "slupkg-ms"
                Return "application/x-ms-license"
            Case "smd"
                Return "audio/x-smd"
            Case "smi"
                Return "application/octet-stream"
            Case "smx"
                Return "audio/x-smd"
            Case "smz"
                Return "audio/x-smd"
            Case "snd"
                Return "audio/basic"
            Case "snippet"
                Return "application/xml"
            Case "snp"
                Return "application/octet-stream"
            Case "sol"
                Return "text/plain"
            Case "sor"
                Return "text/plain"
            Case "spc"
                Return "application/x-pkcs7-certificates"
            Case "spl"
                Return "application/futuresplash"
            Case "src"
                Return "application/x-wais-source"
            Case "srf"
                Return "text/plain"
            Case "ssisdeploymentmanifest"
                Return "text/xml"
            Case "ssm"
                Return "application/streamingmedia"
            Case "sst"
                Return "application/vnd.ms-pki.certstore"
            Case "stl"
                Return "application/vnd.ms-pki.stl"
            Case "sv4cpio"
                Return "application/x-sv4cpio"
            Case "sv4crc"
                Return "application/x-sv4crc"
            Case "svc"
                Return "application/xml"
            Case "swf"
                Return "application/x-shockwave-flash"
            Case "t"
                Return "application/x-troff"
            Case "tar"
                Return "application/x-tar"
            Case "tcl"
                Return "application/x-tcl"
            Case "testrunconfig"
                Return "application/xml"
            Case "testsettings"
                Return "application/xml"
            Case "tex"
                Return "application/x-tex"
            Case "texi"
                Return "application/x-texinfo"
            Case "texinfo"
                Return "application/x-texinfo"
            Case "tgz"
                Return "application/x-compressed"
            Case "thmx"
                Return "application/vnd.ms-officetheme"
            Case "thn"
                Return "application/octet-stream"
            Case "tif"
                Return "image/tiff"
            Case "tiff"
                Return "image/tiff"
            Case "tlh"
                Return "text/plain"
            Case "tli"
                Return "text/plain"
            Case "toc"
                Return "application/octet-stream"
            Case "tr"
                Return "application/x-troff"
            Case "trm"
                Return "application/x-msterminal"
            Case "trx"
                Return "application/xml"
            Case "ts"
                Return "video/vnd.dlna.mpeg-tts"
            Case "tsv"
                Return "text/tab-separated-values"
            Case "ttf"
                Return "application/octet-stream"
            Case "tts"
                Return "video/vnd.dlna.mpeg-tts"
            Case "txt"
                Return "text/plain"
            Case "u32"
                Return "application/octet-stream"
            Case "uls"
                Return "text/iuls"
            Case "user"
                Return "text/plain"
            Case "ustar"
                Return "application/x-ustar"
            Case "vb"
                Return "text/plain"
            Case "vbdproj"
                Return "text/plain"
            Case "vbk"
                Return "video/mpeg"
            Case "vbproj"
                Return "text/plain"
            Case "vbs"
                Return "text/vbscript"
            Case "vcf"
                Return "text/x-vcard"
            Case "vcproj"
                Return "application/xml"
            Case "vcs"
                Return "text/plain"
            Case "vcxproj"
                Return "application/xml"
            Case "vddproj"
                Return "text/plain"
            Case "vdp"
                Return "text/plain"
            Case "vdproj"
                Return "text/plain"
            Case "vdx"
                Return "application/vnd.ms-visio.viewer"
            Case "vml"
                Return "text/xml"
            Case "vscontent"
                Return "application/xml"
            Case "vsct"
                Return "text/xml"
            Case "vsd"
                Return "application/vnd.visio"
            Case "vsi"
                Return "application/ms-vsi"
            Case "vsix"
                Return "application/vsix"
            Case "vsixlangpack"
                Return "text/xml"
            Case "vsixmanifest"
                Return "text/xml"
            Case "vsmdi"
                Return "application/xml"
            Case "vspscc"
                Return "text/plain"
            Case "vss"
                Return "application/vnd.visio"
            Case "vsscc"
                Return "text/plain"
            Case "vssettings"
                Return "text/xml"
            Case "vssscc"
                Return "text/plain"
            Case "vst"
                Return "application/vnd.visio"
            Case "vstemplate"
                Return "text/xml"
            Case "vsto"
                Return "application/x-ms-vsto"
            Case "vsw"
                Return "application/vnd.visio"
            Case "vsx"
                Return "application/vnd.visio"
            Case "vtx"
                Return "application/vnd.visio"
            Case "wav"
                Return "audio/wav"
            Case "wave"
                Return "audio/wav"
            Case "wax"
                Return "audio/x-ms-wax"
            Case "wbk"
                Return "application/msword"
            Case "wbmp"
                Return "image/vnd.wap.wbmp"
            Case "wcm"
                Return "application/vnd.ms-works"
            Case "wdb"
                Return "application/vnd.ms-works"
            Case "wdp"
                Return "image/vnd.ms-photo"
            Case "webarchive"
                Return "application/x-safari-webarchive"
            Case "webtest"
                Return "application/xml"
            Case "wiq"
                Return "application/xml"
            Case "wiz"
                Return "application/msword"
            Case "wks"
                Return "application/vnd.ms-works"
            Case "wlmp"
                Return "application/wlmoviemaker"
            Case "wlpginstall"
                Return "application/x-wlpg-detect"
            Case "wlpginstall3"
                Return "application/x-wlpg3-detect"
            Case "wm"
                Return "video/x-ms-wm"
            Case "wma"
                Return "audio/x-ms-wma"
            Case "wmd"
                Return "application/x-ms-wmd"
            Case "wmf"
                Return "application/x-msmetafile"
            Case "wml"
                Return "text/vnd.wap.wml"
            Case "wmlc"
                Return "application/vnd.wap.wmlc"
            Case "wmls"
                Return "text/vnd.wap.wmlscript"
            Case "wmlsc"
                Return "application/vnd.wap.wmlscriptc"
            Case "wmp"
                Return "video/x-ms-wmp"
            Case "wmv"
                Return "video/x-ms-wmv"
            Case "wmx"
                Return "video/x-ms-wmx"
            Case "wmz"
                Return "application/x-ms-wmz"
            Case "wpl"
                Return "application/vnd.ms-wpl"
            Case "wps"
                Return "application/vnd.ms-works"
            Case "wri"
                Return "application/x-mswrite"
            Case "wrl"
                Return "x-world/x-vrml"
            Case "wrz"
                Return "x-world/x-vrml"
            Case "wsc"
                Return "text/scriptlet"
            Case "wsdl"
                Return "text/xml"
            Case "wvx"
                Return "video/x-ms-wvx"
            Case "x"
                Return "application/directx"
            Case "xaf"
                Return "x-world/x-vrml"
            Case "xaml"
                Return "application/xaml+xml"
            Case "xap"
                Return "application/x-silverlight-app"
            Case "xbap"
                Return "application/x-ms-xbap"
            Case "xbm"
                Return "image/x-xbitmap"
            Case "xdr"
                Return "text/plain"
            Case "xht"
                Return "application/xhtml+xml"
            Case "xhtml"
                Return "application/xhtml+xml"
            Case "xla"
                Return "application/vnd.ms-excel"
            Case "xlam"
                Return "application/vnd.ms-excel.addin.macroenabled.12"
            Case "xlc"
                Return "application/vnd.ms-excel"
            Case "xld"
                Return "application/vnd.ms-excel"
            Case "xlk"
                Return "application/vnd.ms-excel"
            Case "xll"
                Return "application/vnd.ms-excel"
            Case "xlm"
                Return "application/vnd.ms-excel"
            Case "xls"
                Return "application/vnd.ms-excel"
            Case "xlsb"
                Return "application/vnd.ms-excel.sheet.binary.macroenabled.12"
            Case "xlsm"
                Return "application/vnd.ms-excel.sheet.macroenabled.12"
            Case "xlsx"
                Return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Case "xlt"
                Return "application/vnd.ms-excel"
            Case "xltm"
                Return "application/vnd.ms-excel.template.macroenabled.12"
            Case "xltx"
                Return "application/vnd.openxmlformats-officedocument.spreadsheetml.template"
            Case "xlw"
                Return "application/vnd.ms-excel"
            Case "xml"
                Return "text/xml"
            Case "xmta"
                Return "application/xml"
            Case "xof"
                Return "x-world/x-vrml"
            Case "xoml"
                Return "text/plain"
            Case "xpm"
                Return "image/x-xpixmap"
            Case "xps"
                Return "application/vnd.ms-xpsdocument"
            Case "xrm-ms"
                Return "text/xml"
            Case "xsc"
                Return "application/xml"
            Case "xsd"
                Return "text/xml"
            Case "xsf"
                Return "text/xml"
            Case "xsl"
                Return "text/xml"
            Case "xslt"
                Return "text/xml"
            Case "xsn"
                Return "application/octet-stream"
            Case "xss"
                Return "application/xml"
            Case "xtp"
                Return "application/octet-stream"
            Case "xwd"
                Return "image/x-xwindowdump"
            Case "z"
                Return "application/x-compress"
            Case "zip"
                Return "application/x-zip-compressed"
            Case Else
                Return "application/octet-stream"
        End Select
    End Function

End Class
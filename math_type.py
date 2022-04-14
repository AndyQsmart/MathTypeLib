import platform
from os import path
from ctypes import c_char_p, create_string_buffer

windll = None
cdll = None
system_type = platform.system()
if system_type == 'Windows':
    from ctypes import windll as ctypes_windll
    windll = ctypes_windll
elif system_type == 'Darwin':
    from ctypes import cdll as ctypes_cdll
    cdll = ctypes_cdll

# Types
# 'Picture specifier
# Public Type MTAPI_PICT
#     mm    As Long
#     xExt  As Long
#     yExt  As Long
#     hMF   As Long
# End Type

# Public Type RECT
#     left     As Long
#     top      As Long
#     right    As Long
#     bottom   As Long
# End Type

# ' Picture dimensions
# Public Type MTAPI_DIMS
#     baseline As Integer  ' distance of baseline from bottom (points)
#     bounds   As RECT     ' bounding rectangle (points)
# End Type

class MathTypeReturnValue:
    mtOK = 0 # no error
    mtNOT_FOUND = -1 # trouble finding the MT erver, usually indicates a bad session ID
    mtCANT_RUN = -2 # could not start the MT server
    mtBAD_VERSION = -3 # server / DLL version mismatch
    mtIN_USE = -4 # server is busy/in-use
    mtNOT_RUNNING = -5 # server aborted
    mtRUN_TIMEOUT = -6 # connection to the server failed due to time-out 
    mtNOT_EQUATION = -7 # an API call that expects an equation could not find one
    mtFILE_NOT_FOUND = -8 # a preference, translator or other file could not be found
    mtMEMORY = -9 # a buffer too small to hold the result of an API call was passed in
    mtBAD_FILE = -10 # file found was not a translator
    mtDATA_NOT_FOUND = -11 # unable to read preferences from MTEF on the clipboard
    mtTOO_MANY_SESSIONS = -12 # too many open connections to the SDK
    mtSUBSTITUTION_ERROR = -13 # problem with substition error during a call to MTXFormEqn
    mtTRANSLATOR_ERROR = -14 # there was an error in compiling or in execution of a translator 
    mtPREFERENCE_ERROR = -15 # could not set preferences
    mtBAD_PATH = -16 # a bad path was encountered when trying to write to a file
    mtFILE_ACCESS = -17 # a file could not be written to
    mtFILE_WRITE_ERROR = -18 # a file could not be written to
    mtBAD_DATA = -19 # (deprecated)
    mtERROR = -9999 # other error

class MTXFormSetTranslatorOptions:
    mtxfmTRANSL_INC_NONE = 0
    mtxfmTRANSL_INC_NAME = 1
    mtxfmTRANSL_INC_DATA = 2
    mtxfmTRANSL_INC_MTDEFAULT = 4
    # append 'Z' to text equations placed on clipboard
    # kludge to fix Word's trailing CRLF truncation bug
    mtxfmTRANSL_INC_CLIPBOARD_EXTRA = 8

class MTGetTranslatorsInfoIndex:
    mttrnCOUNT = 1
    mttrnMAX_NAME = 2
    mttrnMAX_DESC = 3
    mttrnMAX_FILE = 4
    mttrnOPTIONS = 5

class MTXFormSetPrefsType:
    mtxfmPREF_EXISTING = 1
    mtxfmPREF_MTDEFAULT = 2
    mtxfmPREF_USER = 3

class MTXFormGetStatusIndex:
    mtxfmSTAT_ACTUAL_LEN = -1
    mtxfmSTAT_TRANSL = -2
    mtxfmSTAT_PREF = -3

class MTXFormEqnSrc:
    mtxfmPREVIOUS = -1
    mtxfmCLIPBOARD = -2
    mtxfmLOCAL = -3
    mtxfmFILE = -4

class MTXFormEqnDst(MTXFormEqnSrc):
    pass

class MTXFormEqnSrcFmt:
    mtxfmMTEF = 4
    mtxfmHMTEF = 5
    mtxfmPICT = 6
    mtxfmTEXT = 7
    mtxfmHTEXT = 8
    mtxfmGIF = 9
    mtxfmEPS_NONE = 10
    mtxfmEPS_WMF = 11
    mtxfmEPS_TIFF = 12

class MTXFormEqnDstFmt(MTXFormEqnSrcFmt):
    pass

class MTSetMTPrefsMode:
    mtprfMODE_NEXT_EQN = 1
    mtprfMODE_MTDEFAULT = 2
    mtprfMODE_INLINE = 4

class MTXFormAddVarSubOptions:
    mtxfmSUBST_ALL = 0
    mtxfmSUBST_ONE = 1

class MTXFormAddVarSubFindType:
    mtxfmVAR_SUB_PLAIN_TEXT = 0
    mtxfmVAR_SUB_MTEF_TEXT = 1
    mtxfmVAR_SUB_MTEF_BINARY = 2
    mtxfmVAR_SUB_DELETE = 3

class MTXFormAddVarSubReplaceType(MTXFormAddVarSubFindType):
    pass

class MTXFormAddVarSubReplaceStyle:
    mtxfmSTYLE_TEXT = 1
    mtxfmSTYLE_FUNCTION = 2
    mtxfmSTYLE_VARIABLE = 3
    mtxfmSTYLE_LCGREEK = 4
    mtxfmSTYLE_UCGREEK = 5
    mtxfmSTYLE_SYMBOL = 6
    mtxfmSTYLE_VECTOR = 7
    mtxfmSTYLE_NUMBER = 8

class MTAPIConnectOptions:
    mtinitLAUNCH_AS_NEEDED = 0
    mtinitLAUNCH_NOW = 1

class MTEquationOnClipboardReturnValue:
    mtOLE_EQUATION = 1 # equation OLE 1.0 object on clipboard
    mtWMF_EQUATION = 2 # Windows metafile equation graphic (not OLE object) on clipboard
    mtMAC_PICT_EQUATION = 4 # Macintosh PICT equation graphic (not OLE object) on clipboard
    mtOLE2_EQUATION = 8 # equation OLE 2.0 object on clipboard

class MTGetLastDimensionIndex:
    mtdimWIDTH = 1
    mtdimHEIGHT = 2
    mtdimBASELINE = 3
    mtdimHORIZ_POS_TYPE = 4
    mtdimHORIZ_POS = 5

class EnumTranslatorsReturnValue:
    def __init__(self, transName, transDesc, transFile, next, status):
        self.transName = transName
        self.transDesc = transDesc
        self.transFile = transFile
        self.next = next
        self.status = status

class MTEFData:
    def __init__(self, byte_data):
        self._byte_data = byte_data

    def getBytes(self):
        return self._byte_data

    @classmethod
    def fromWmf(self, wmf_data):
        # 临时性读取方案
        hex_str = wmf_data.hex()
        index = 0

        dsmt_str = 'DSMT'.encode().hex()
        mtef_index = -1
        # 查询起始字符
        while index < len(hex_str):
            is_dsmt_index = True
            s_index = 0
            while s_index < len(dsmt_str):
                if dsmt_str[s_index] != hex_str[index+s_index]:
                    is_dsmt_index = False
                    break
                s_index += 1

            if is_dsmt_index and mtef_index == -1:
                mtef_index = index-10
                break
            
            index += 1

        end_str = '0000000000'
        end_index = len(hex_str)
        
        mtef_str = ''
        index = mtef_index
        while index < len(hex_str) and index < end_index:
            mtef_str += hex_str[index]
            index += 1

        return MTEFData(bytes.fromhex(mtef_str))

class MathTypeLib:
    def __init__(self):
        self.mt_lib = None

    def LoadLibrary(self):
        if system_type == 'Windows':
            lib_path = path.join(path.dirname(__file__), 'MT6.dll')
            self.mt_lib = windll.LoadLibrary(lib_path)
        elif system_type == 'Darwin':
            # cdll.LoadLibrary('/Library/Frameworks/MT6Lib.framework/MT6Lib')
            pass
        # GetTranslatorsInfo数据缓存
        self.translator_count = None
        self.translator_name_max_len = None
        self.translator_description_max_len = None
        self.translator_file_name_max_len = None

    def MTAPIVersion(self, api):
        """
            函数说明：
                Public Function MTAPIVersion (
                    api As Integer
                ) As Long
                Gets version info for a particular API set.

            参数说明：
                api: the value of the API that's required (only api = 5 is allowed currently and its version is 5.2)

            返回值说明：
                Return value
                Returns version or mpMTDLL_NOT_FOUND. Version is 0 if API set unknown or hi-byte (of lo-word) = major version, lo-byte (of lo-word) = minor version.

            举例：
                self.MTAPIVersion(5)
        """
        return self.mt_lib.MTAPIVersion(api)

    def MTAPIConnect(self, options, timeout):
        """
        函数说明：
            Public Function MTAPIConnect (
                 options As Integer,
                 timeout As Integer 
            ) As Long
            This function initializes the API (prepares for use). This function should always be called before any other API function (except MTAPIVersion). Other API functions will return an error if this function has not been called. Calls to MTAPIConnect should always be paired with calls to MTAPIDisconnect, and they may not be nested. Prior to calling MTAPIDisconnect subsequent calls to MTAPIConnect will do nothing and just return mtOK. This function may not be called from MathPage.WLL, use the MathPage.WLL function MTInitAPI instead. Note: This function was formerly named MTInitAPI, and calling it was optional (i.e. if a given call to another API function detected that MTInitAPI had not been called it would call it for you). This is not the case with MTAPIConnect. You must call MTAPIConnect prior to any other API calls (except MTAPIVersion).

        参数说明：
            options: mtinitLAUNCH_NOW => launch MathType server immediately
                    mtinitLAUNCH_AS_NEEDED => launch MathType when first needed
            timeout: # of seconds to wait before timing out when attempting to launch MathType. If timeOut = -1 then will never timeout.  This value is eventually passed to RPCConnectToServer where it is not currently used

        返回值说明：
            Return value
            Returns mtOK, mtMT_NOT_FOUND, mtMT_CANT_RUN, mtMT_BAD_VERSION, mtMT_IN_USE, mtMT_RUN_TIME_OUT, OR mtERROR.

        举例：
            self.MTAPIConnect(MTAPIConnectOptions.mtinitLAUNCH_AS_NEEDED, 30)
        """
        return self.mt_lib.MTAPIConnect(options, timeout)

    def APIConnect(self, options, timeout):
        return self.MTAPIConnect(options, timeout)

    def MTAPIDisconnect(self):
        """
        函数说明：
            Public Function MTAPIDisconnect () As Long
            This function must always be called after MTAPIConnect. It closes connection with MathType and terminates the API (no further usage). This function may not be called from MathPage.WLL, use the MathPage.WLL function MTTermAPI instead.

        返回值说明：
            Return value
            Returns mtOK.

        举例：
            self.MTAPIDisconnect()
        """
        return self.mt_lib.MTAPIDisconnect()

    def APIDisconnect(self):
        return self.MTAPIDisconnect()

    def MTEquationOnClipboard(self):
        """
        函数说明：
            Public Function MTEquationOnClipboard () As Long
            Check for the type of equation on the clipboard, if any.

        返回值说明：
            Return value
            If equation on the clipboard, returns type of eqn data:
                mtOLE_EQUATION, mtWMF_EQUATION, mtMAC_PICT_EQUATION,
            Otherwise status value:
                mtNOT_EQUATION - no eqn on clipboard
                mtMEMORY - insufficient memory for prog ID
                mtERROR - any other error

        举例：
            self.MTEquationOnClipboard()
        """
        return self.mt_lib.MTEquationOnClipboard()

    def EquationOnClipboard(self):
        return self.MTEquationOnClipboard()

    def MTClearClipboard(self):
        """
        函数说明：
            Public Function MTClearClipboard () As Long
            Clears the clipboard contents.

        返回值说明：
           Return value
            Returns mtOK.

        举例：
            self.MTClearClipboard()
        """
        return self.mt_lib.MTClearClipboard()

    def ClearClipboard(self):
        return self.MTClearClipboard()

    def MTGetLastDimension(self, dimIndex):
        """
        函数说明：
            Public Function MTGetLastDimension (
                dimIndex As Integer 
            ) As Long
            Gets a dimension from the last equation copied to the clipboard or written to a file.

        参数说明：
            dimIndex: desired dimension. One of the following:
                mtdimWIDTH, mtdimHEIGHT, mtdimBASELINE, mtdimHORIZ_POS_TYPE, mtdimHORIZ_POS

        返回值说明：
            Return value
            If successful (>0), value of desired dimension in 32nds of a point.
            Otherwise, error status:
                mtNOT_EQUATION - no equation to take dimension from
                mtERROR - bad value for dimIndex

        举例：
            self.MTGetLastDimension(MTGetLastDimensionIndex.mtdimWIDTH)
        """
        return self.mt_lib.MTGetLastDimension(dimIndex)

    def MTOpenFileDialog(self, fileType, title, dir, file, fileLen):
        """
        函数说明：
            Public Function MTOpenFileDialog (
                fileType As Integer,
                title As String,
                dir As String,
                file As String,
                fileLen As Integer
            ) As Long
            Puts up an open file dialog (Win32 only). Calls GetForegroundWindow for parent, upon which it gets centered.

        参数说明：
            fileType: 1 for MT preference files
            title: dialog window title
            dir: default directory (may be empty or NULL)
            file: result: new filename
            fileLen: maximum number of characters in filename

        返回值说明：
            Return value
            Returns 1 for OK, 0 for Cancel

        举例：
            self.MTOpenFileDialog(1, 'test', None, 'test', 32)
        """
        self.mt_lib.MTOpenFileDialog(fileType, title, dir, file, fileLen)

    def MTGetPrefsFromClipboard(self, prefs, prefsLen):
        """
        函数说明：
            Public Function MTGetPrefsFromClipboard (
                prefs As String,
                prefsLen As Integer
            ) As Long
            Gets equation preferences from the MathType equation currently on the clipboard.

        参数说明：
            prefs: [out] Preference string (if sizeStr > 0)
            prefsLen: [in] Size of prefStr (inc. null) or 0

        返回值说明：
            Return value
            If prefsLen = 0 then this is the size required for prefs.
            Otherwise it's a status:
                mtOK Success
                mtMEMORY Not enough memory for to store preferences
                mtNOT_EQUATION Not equation on clipboard
                mtBAD_VERSION No preference data found in equation
                mtERROR Other error

        举例：
        """
        return self.mt_lib.MTGetPrefsFromClipboard(prefs, prefsLen)

    def MTGetPrefsFromFile(self, prefFile, prefs, prefsLen):
        """
        函数说明：
            Public Function MTGetPrefsFromFile (
                prefFile As String,
                prefs As String,
                prefsLen As Integer
            ) As Long
            Get equation preferences from the specified preferences file.

        参数说明：
            prefFile: [in] Pathname for the preference file
            prefs: [out] Preference string (if sizeStr > 0)
            prefsLen: [in] Size of prefStr or 0

        返回值说明：
            Return value
            If sizeStr = 0 then this is the size required for prefStr. Otherwise it's a status:
                mtOK Success
                mtMEMORY Not enough memory for to store preferences
                mtFile_NOT_FOUND File does not exist or bad pathname
                mtERROR Other error

        举例：
        """
        return self.mt_lib.MTGetPrefsFromFile(prefFile, prefs, prefsLen)

    def MTConvertPrefsToUIForm(self, inPrefs, outPrefs, outPrefsLen):
        """
        函数说明：
            Public Function MTConvertPrefsToUIForm (
                inPrefs As String,
                outPrefs As String,
                outPrefsLen As Integer
            ) As Long
            Convert internal preferences string to a form to be presented to the user.

        参数说明：
            inPrefs: [in] internal preferences string
            outPrefs: [out] Preference string (if sizeStr > 0)
            outPrefsLen: [in] Size of outPrefStr (inc. null) or 0 to get length

        返回值说明：
            Return value
            If outPrefsLen = 0 then this is the size required for outPrefStr, else it's a status:
                mtOK Success
                mtMEMORY Not enough memory for to store preferences
                mtERROR Other error

        举例：
        """
        return self.mt_lib.MTConvertPrefsToUIForm(inPrefs, outPrefs, outPrefsLen)

    def MTGetPrefsMTDefault(self, prefs, prefsLen):
        """
        函数说明：
            Public Function MTGetPrefsMTDefault (
                prefs As String,
                prefsLen As Integer
            ) As Long
            Get MathType's current default equation preferences

        参数说明：
            prefs: [out] Preference string (if sizeStr > 0)
            prefsLen: [in] Size of prefStr or 0

        返回值说明：
            Return value
            If prefsLen = 0 then this is the size required for prefs, otherwise it's a status.
                mtOK Success
                mtMEMORY Not enough memory for to store preferences
                mtERROR Other error

        举例：
        """
        return self.mt_lib.MTGetPrefsMTDefault(prefs, prefsLen)

    def MTSetMTPrefs(self, mode, prefs, timeout):
        """
        函数说明：
            Public Function MTSetMTPrefs (
                mode As Integer,
                prefs As String,
                timeout As Integer
            ) As Long
            Set MathType's default preferences for new equations.

        参数说明：
            mode: [in] Specifies the way the preferences will be applied
                mtprfMODE_NEXT_EQN => Apply to next new equation (see timeOut)
                mtprfMODE_MTDEFAULT => Set MathType's defaults for new equations
                mtprfMODE_INLINE => makes next eqn inline
            prefs: [in] Null terminated preference string
            timeout: [in] Number of seconds to wait for new equation (used only when mode = 1), Note: -1 means wait forever

        返回值说明：
            Return value
            Returns status:
                mtOK Success
                mtBAD_DATA Bad prefs string,
                mtERROR Any other error

        举例：
        """
        return self.mt_lib.MTSetMTPrefs(mode, prefs, timeout)

    def MTGetTranslatorsInfo(self, infoIndex):
        """
        函数说明：
            Public Function MTGetTranslatorsInfo (
                 infoIndex As Integer
            ) As Long
            Get information about the current set of translators.

        参数说明：
            infoindex: [in] A flag indicating what info to return:
                1 mttrnCOUNT => Total number of translators
                2 mttrnMAX_NAME => Maximum size of any translator name
                3 mttrnMAX_DESC => Maximum size of any translator description string
                4 mttrnMAX_FILE => Maximum size of any translator file name
                5 mttrnOPTIONS => Translator options

        返回值说明：
            Return value
            If >= 0 then this value is the information specified by infoIndex, otherwise its a status:
                mtERROR Bad value for infoIndex

        举例：
            self.MTGetTranslatorsInfo(MTGetTranslatorsInfoIndex.mttrnCOUNT)
        """
        return self.mt_lib.MTGetTranslatorsInfo(infoIndex)

    def GetTranslatorsInfo(self, infoIndex):
        return self.MTGetTranslatorsInfo(infoIndex)

    def MTEnumTranslators(self, index, transName, transNameLen, transDesc, transDescLen, transFile, transFileLen):
        """
        函数说明：
            Public Function MTEnumTranslators (
                index As Integer,
                transName As String,
                transNameLen As Integer,
                transDesc As String,
                transDescLen As Integer,
                transFile As String,
                transFileLen As Integer
            ) As Long
            Enumerate the available equation (TeX, etc.) translators.

        参数说明：
            index: [in] Index of the translator to enumerate (must be initialized to 1 by the caller)
            transName: [out] Translator name
            transNameLen: [in] Size of tShort. (May be set to zero)
            transDesc: [out] Translator descriptor string
            transDescLen: [in] Size of transDesc. (May be set to zero)
            transFile: [out] Translator file name
            transFileLen: [in] Size of transFile. (May be set to zero)

        返回值说明：
            Return value
            If > 0 then this value is the index of next translator to enumerate, (i.e. the caller should pass this value in for index).
            Otherwise, a status:
                mtOK Success (no more translators in the list)
                mtMEMORY Not enough room in transName, transDesc, or transFile
                mtERROR Any other failure

        举例：
                transName = create_string_buffer(38)
                transDesc = create_string_buffer(79)
                transFile = create_string_buffer(30)
                print(self.MTEnumTranslators(1, transName, 38, transDesc, 79, transFile, 30))
                print('transName:', transName.value)
                print('transDesc:', transDesc.value)
                print('transFile:', transFile.value)
        """
        return self.mt_lib.MTEnumTranslators(index, transName, transNameLen, transDesc, transDescLen, transFile, transFileLen)

    def EnumTranslators(self, index):
        if self.translator_count is None or self.translator_name_max_len is None or self.translator_description_max_len is None or self.translator_file_name_max_len is None:
            self.translator_count = self.GetTranslatorsInfo(MTGetTranslatorsInfoIndex.mttrnCOUNT)
            self.translator_name_max_len = self.GetTranslatorsInfo(MTGetTranslatorsInfoIndex.mttrnMAX_NAME)
            self.translator_description_max_len = self.GetTranslatorsInfo(MTGetTranslatorsInfoIndex.mttrnMAX_DESC)
            self.translator_file_name_max_len = self.GetTranslatorsInfo(MTGetTranslatorsInfoIndex.mttrnMAX_FILE)

        transName = create_string_buffer(self.translator_name_max_len)
        transDesc = create_string_buffer(self.translator_description_max_len)
        transFile = create_string_buffer(self.translator_file_name_max_len)
        next_index = self.MTEnumTranslators(
            index,
            transName, self.translator_name_max_len,
            transDesc, self.translator_description_max_len,
            transFile, self.translator_file_name_max_len 
        )

        return EnumTranslatorsReturnValue(
            transName.value.decode() ,
            transDesc.value.decode(),
            transFile.value.decode(),
            next_index, next_index
        )

    def MTXFormReset(self):
        """
        函数说明：
            Public Function MTXFormReset () As Long
            Resets to default options for MTXFormEqn (i.e. no substitutions, no translation, and use existing preferences).

        返回值说明：
            Return value
            Return mtOK.

        举例：
        """
        return self.mt_lib.MTXFormReset()

    def MTXFormAddVarSub(self, options, findType, find, findLen, replaceType, replace, replaceLen, replaceStyle):
        """
        函数说明：
            Public Function MTXFormAddVarSub (
                options As Integer,
                findType As Integer,
                find As String,
                findLen As Long,
                replaceType As Integer,
                replace As String,
                replaceLen As Long,
                replaceStyle As Integer
            ) As Long
            Adds a variable substitution to be performed with next MTXFormEqn (may be called 0 or more times).

        参数说明：
            options: mtxfmSUBST_ALL or mtxfmSUBST_ONE
            findType: type of data in find arg (must be mtxfmVAR_SUB_PLAIN_TEXT for now)
            find: equation text to be found and replaced (null-terminated text string for now)
            findLen:	length of find arg data (ignored for now)
            replaceType: type of data in replace arg:
                mtxfmVAR_SUB_PLAIN_TEXT
                mtxfmVAR_SUB_MTEF_TEXT
                mtxfmVAR_SUB_MTEF_BINARY
                mtxfmVAR_SUB_DELETE - delete the "find" text
            replace: equation text to replace "find" arg with
            replaceLen: if replaceType = mtxfmVAR_SUB_MTEF_BINARY, length of replace arg data
            replaceStyle: if replaceType = mtxfmVAR_SUB_PLAIN_TEXT, style of replacement text:
                mtxfmSTYLE_TEXT mtxfmSTYLE_FUNCTION
                mtxfmSTYLE_VARIABLE
                mtxfmSTYLE_LCGREEK
                mtxfmSTYLE_UCGREEK
                mtxfmSTYLE_SYMBOL
                mtxfmSTYLE_VECTOR
                mtxfmSTYLE_NUMBER

        返回值说明：
            Return value
            Returns status:
                mtOK - success
                mtERROR - some other error

        举例：
        """
        return self.mt_lib.MTXFormAddVarSub(options, findType, find, findLen, replaceType, replace, replaceLen, replaceStyle)

    def MTXFormSetTranslator(self, options, transName):
        """
        函数说明：
            Public Function MTXFormSetTranslator (
                options As Integer,
                transName As String
            ) As Long
            Specify translation to be performed with the next MTXFormEqn

        参数说明：
            options: [in] One or more (OR'd together) of:
                mtxfmTRANSL_INC_NONE - no options
                mtxfmTRANSL_INC_NAME - include the translator's name in translator output
                mtxfmTRANSL_INC_DATA - include MathType equation data in translator output
                mtxfmTRANSL_INC_MTDEFAULT use MathType's defaults
                mtxfmTRANSL_INC_CLIPBOARD_EXTRA
            transName: [in] File name of translator to be used, NULL for no translation

        返回值说明：
            Return value
            Returns status:
                mtOK - success
                mtFILE_NOT_FOUND - could not find translator
                mtTRANSLATOR_ERROR - errors compiling translator
                mtERROR - some other error

        举例：
            self.MTXFormSetTranslator(MTXFormSetTranslatorOptions.mtxfmTRANSL_INC_MTDEFAULT, "MathML2 (namespace attr).tdl".encode('utf-8'))  # mtxfmTRANSL_INC_MTDEFAULT
        """
        return self.mt_lib.MTXFormSetTranslator(options, transName)

    def XFormSetTranslator(self, options, transName):
        return self.MTXFormSetTranslator(options, transName)

    def MTXFormSetPrefs(self, prefType, prefStr):
        """
        函数说明：
            Public Function MTXFormSetPrefs (
                prefType As Integer,
                prefStr As String
            ) As Long
            Specify a new set of preferences to be used with the next MTXFormEqn.

        参数说明：
            prefType: [in] One of the following:
                mtxfmPREF_EXISTING - use existing preferences
                mtxfmPREF_MTDEFAULT - use MathType's default preferences
                mtxfmPREF_USER - use specified preferences
            prefStr: [in] Preferences to apply (mtxfmPREF_USER)

        返回值说明：
            Return value
            Returns mtOK or mtERROR.

        举例：
        """
        return self.mt_lib.MTXFormSetPrefs(prefType, prefStr)

    def MTXFormEqn(self, src, srcFmt, srcData, srcDataLen, dst, dstFmt, dstData, dstDataLen, dstPath, dims):
        """
        函数说明：
            Public Function MTXFormEqn (
                src As Integer,
                srcFmt As Integer,
                srcData As String,
                srcDataLen As Long,
                dst As Integer,
                dstFmt As Integer,
                dstData As String,
                dstDataLen As Long,
                dstPath As String,
                dims As MTAPI_DIMS
            ) As Long
            Transforms an equation (uses options specified via MTXFormAddVarSub, MTXFormSetTranslator, and MTXFormSetPrefs)
            Note: Variations involving mtxfm_PICT or dstFmt=mtxfm_HMTEF are not callable from VBA.

        参数说明：
            src: [in] Equation data source, one of:
                mtxfmPREVIOUS => data from previous result
                mtxfmCLIPBOARD => data on clipboard
                mtxfmLOCAL => data passed (i.e. in srcData)
            srcFmt: [in] Equation source data format, one of:
                mtxfmMTEF
                mtxfmPICT
                mtxfmTEXT
                Note: srcFmt, srcData, and srcDataLen are used only if src is mtxfmLOCAL
            srcData: [in] Depends on data source (srcFmt)
                if srcFmt is mtxfmMTEF, then srcData must point to MTEF-binary (BYTE *) data
                if srcFmt is mtxfmPICT, then srcData must point to PICT (MTAPI_PICT *) data
                if srcFmt is mtxfmTEXT, then srcData must point to either MTEF-text or plain text (CHAR *) data
            srcDataLen: [in] # of bytes in srcData
            dst: [in] Equation data destination, one of:
                mtxfmCLIPBOARD => transformed data placed on clipboard
                mtxfmFILE => transformed data in the file specified by dstPath
                mtxfmLOCAL => transformed data in dstData
            dstFmt: [in] Equation data format, one of:
                mtxfmMTEF
                mtxfmHMTEF
                mtxfmPICT
                mtxfmGIF
                mtxfmTEXT
                mtxfm HTEXT
                Note: dstFmt, dstData, and dstDataLen are used only if dst is mtxfmLOCAL (The data is placed on the clipboard in either an OLE object or translator text) or mtxfmFILE (dstFmt specifies the file format, dstPath specifies the name of the file to create).
            dstData: [out] Depends on data destination (dstFmt)
                if dstFmt is mtxfmMTEF, then dstData points to MTEF-binary data
                if dstFmt is mtxfmHMTEF, then dstData is a handle to MTEF-binary data
                if dstFmt is mtxfmPICT, then dstData points to PICT data
                if dstFmt is mtxfmGIF, then dstData points to GIF data
                if dstFmt is mtxfmTEXT, then dstData points to translated text or, if no translator, MTEF-text data
                if dstFmt is mtxfmHTEXT, then dstData is a handle to translated text or, if no translator, MTEF-text data
                Note: If a translator was specified, dstFmt must be either mtxfmTEXT or mtxfmHTEXT for the translation to be performed.
            dstDataLen: [in] # of bytes in dstData (used for mtxfmLOCAL only)
            dstPath: [in] destination pathname (used if dst == mtxfmFILE only, may be NULL if not used)
            dims: [out] pict dimensions, may be NULL (valid only for dst = mtxfmPICT). See MTAPI_DIMS definition above in Types section.
            Notes:
                if src is mtfxmLOCAL, then srcFmt, srcData, and srcDataLen must be supplied
                if dst is mtfxmLOCAL, then dstFmt, dstData, and dstDataLen must be supplied
                one can convert from MTEF text to various file types by passing these parameters:
                src: mtxfmCLIPBOARD
                srcFmt: mtxfmTEXT
                dst: mtxfmFILE
                dstFmt: one of: mtxfmPICT (wmf), mtxfmGIF (gif), mtxfmEPS_NONE (eps), mtxfmEPS_WMF (eps), mtxfmEPS_TIFF (eps)
                dstPath: (path and file name for output file)
                all other parameters are ignored
                The SDK sample application, ConvertEquations, demonstrates how to call MTXFormEqn for each of the supported combinations of source format and type, and destination format and type. The table in note #5 below also shows all of the supported combinations.
                The following table shows the various supported conversion options in terms of where the data is to be converted from/to and what formats are supported: 
           	    |           |                           |                                                               Destination                                                             |
                |           |                           |                #clipboard#                 |	                    #file#                 |                 #local#                    |
                |  #Source# |         #format#          | text | MTEF, EPS, or GIF | WMF | Emb. Obj. | text or MTEF | EPS, GIF, or WMF | Emb. Obj. | text or MTEF | EPS, GIF, WMF, or Emb. Obj. |
                | clipboard |       text or MTEF        | Yes  |                   | Yes |           |              |       Yes        |	 	   |      Yes     |                             |
                |           |     EPS, GIF, or WMF      |      |                   |     |           |              |                  |           |              |                             |
                |           |         Emb. Obj.         | Yes  |                   | Yes |           |              |       Yes	 	   |           |      Yes     |                             |
                |   file	|       text or MTEF        |      |                   |     |           |              |                  |           |              |                             |
                |           |     EPS, GIF, or WMF      | Yes  |                   | Yes |           |              |       Yes	 	   |           |      Yes     |                             |
                |           |         Emb. Obj.         |      |                   |     |           |              |                  |           |              |                             |
                |   local	|       text or MTEF        | Yes  |                   |     |           |              |       Yes	 	   |           |      Yes     |                             |
                |           |EPS, GIF, WMF, or Emb. Obj.|      |                   |     |           |              |                  |           |              |                             |

        返回值说明：
            Return value
            Returns status:
                mtOK - success
                mtNOT_EQUATION - source data does not contain MTEF
                mtSUBSTITUTION_ERROR - could not perform one or more subs
                mtTRANSLATOR_ERROR - errors occured during translation (translation not done)
                mtPREFERENCE_ERROR - could not set perferences
                mtMEMORY - not enough space in dstData
                mtERROR - some other error

        举例：
            self.MTXFormEqn(MTXFormEqnSrc.mtxfmCLIPBOARD, MTXFormEqnSrcFmt.mtxfmPICT, None, 0,
                MTXFormEqnDst.mtxfmCLIPBOARD, MTXFormEqnDstFmt.mtxfmTEXT, None, 0,
                '', None
            )
        """
        return self.mt_lib.MTXFormEqn(src, srcFmt, srcData, srcDataLen, dst, dstFmt, dstData, dstDataLen, dstPath, dims)

    def XFormEqnFromWmf(self, wmf_data, dstFmt=MTXFormEqnDstFmt.mtxfmTEXT):
        """
            dsmFmt仅支持text or MTEF
        """
        mtef_data = MTEFData.fromWmf(wmf_data).getBytes()
        src_data = c_char_p(mtef_data)
        dst_data = create_string_buffer(32*1024) # MTEF最大32Kb,而text格式待定
        res = self.MTXFormEqn(
            MTXFormEqnSrc.mtxfmLOCAL, MTXFormEqnSrcFmt.mtxfmMTEF, src_data, len(mtef_data),
            MTXFormEqnDst.mtxfmLOCAL, dstFmt, dst_data, 32*1024,
            '', None
        )

        if res != MathTypeReturnValue.mtOK:
            reason = ''
            if res == MathTypeReturnValue.mtNOT_EQUATION:
                reason = 'source data does not contain MTEF'
            elif res == MathTypeReturnValue.mtSUBSTITUTION_ERROR:
                reason = 'could not perform one or more subs'
            elif res == MathTypeReturnValue.mtTRANSLATOR_ERROR:
                reason = 'errors occured during translation (translation not done)'
            elif res == MathTypeReturnValue.mtPREFERENCE_ERROR:
                reason = 'could not set perferences'
            elif res == MathTypeReturnValue.mtMEMORY:
                reason = 'not enough space in dstData'
            elif res == MathTypeReturnValue.mtERROR:
                reason = 'some other error'
            # print('MathTypeLib.XFormEqnFromWmf.Error:', reason)
            raise Exception(res, reason)

        return dst_data.value

    def MTXFormGetStatus(self, index):
        """
        函数说明：
            Public Function MTXFormGetStatus (
                index As Integer
            ) As Long
            Check error/status after MTXFormEqn.

        参数说明：
            index: [in] which status to get; described below:
                mtxfmSTAT_PREF, status for set preferences
                mtxfmSTAT_TRANSL, status for translation
                mtxfmSTAT_ACTUAL_LEN
                ≥1, status of the i=th (i=index) variable substitution

        返回值说明：
            Return value
            Depends on the value of 'index', as follows:
            If index = mtxfmSTAT_PREF, status for set preferences:
                mtOK - success setting preferences
                mtBAD_DATA - bad preference data
            If index = mtxfmSTAT_TRANSL, status for translation:
                mtOK - successful translation
                mtFILE_NOT_FOUND - could not find translator
                mtBAD_FILE - file found was not a translator
            If index = mtxfmSTAT_ACTUAL_LEN, number of bytes of data actually returned in dstData (if MTXformEqn succeeded) or the number of bytes required (if MTXformEqn returned
                mtMEMORY - not enough memory was specified for dstData), otherwise 0L.
            If index ≥1, status of the i-th (i = index) variable substitution, either # of times the substitution was performed, or, if < 0, an error status.
            NOTE: returns mtERROR for bad values of index

        举例：
        """
        return self.mt_lib.MTXFormGetStatus(index)

    def MTPreviewDialog(self, parent, title, prefs, closeBtnText, helpBtnText, helpID, helpFile):
        """
        函数说明：
            Public Function MTPreviewDialog (
                parent As Long,
                title As String,
                prefs As String,
                closeBtnText As String,
                helpBtnText As String,
                helpID As Long,
                helpFile As String
            ) As Long
            Puts up a preview dialog for displaying preferences

        参数说明：
            parent: parent window
            title: dialog title
            prefs: text to preview
            closeBtnText: text for Close button (can be NULL for English)
            helpBtnText: text for Help button (can be NULL for English)
            helpID: help topic ID
            helpFile: help file

        返回值说明：
            Return value
            Returns 0 if successful, non-zero if error.

        举例：
        """
        return self.mt_lib.MTPreviewDialog(parent, title, prefs, closeBtnText, helpBtnText, helpID, helpFile)

    def MTGetPathToMathType(self, path, pathMax):
        """
        函数说明：
            Public Function MTGetPathToMathType (
                path As String,
                pathMax As Long
            ) As Long
            Get the path to the MathType folder.

        参数说明：
            path: [out] Fully qualified path to the MathType folder
            pahMax: [in/out] Max size of path on input; actual length on output

        返回值说明：
            Return value
            Returns mtOK on success, otherwise mtERROR.

        举例：
        """
        return self.mt_lib.MTGetPathToMathType(path, pathMax)

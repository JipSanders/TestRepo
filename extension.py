"""Helper routines for EXTENSION command"""

#Licensed Materials - Property of IBM
#IBM SPSS Products: Statistics General
#(c) Copyright IBM Corp. 2005, 2015
#US Government Users Restricted Rights - Use, duplication or disclosure restricted by GSA ADP Schedule Contract with IBM Corp.

# History
# 12-Oct-2007 JKP Initial version
# 14-Nov-2007 Changed to handle Unicode mode
# 21-Feb-2008 Correct error msg for int, float out of range
# 27-Jun-2008  Add checkrequiredparams function
# 12-Jul -2008  Add processcmd function to handle general execution pattern
# 08-aug-2008 Eliminate dependence on spssaux module to improve script efficiency
# 12-aug-2008 Provide for optional VariableDict object in processcmd
# 26-dec-2008 floatex fix for comma decimal
# 10-dec-2009 enable translation, convert exception  messages to pivot tables automatically
# 09-mar-2010 more robust mapping of exception info to pivot table string (handle exceptions returning classes)
# 30-mar-2010 still more robust exception handling
# 01-jun-2010  add code to address language identifier discrepancy for S and T Chinese between SPSS and Python
# 06-jul-2010  fix premature use of _ function
# 27-dec-2010 add environment-variable control over whether exceptions are converted to pivot tables
#                     or reraised (for debugging purposes)
# 20-dec-2012 replace inspect.stack call with limited version due to Python/win64/Windows 8 bug
# 26-aug-2014 add helper function for extension command help
# 19-aug-2015 add fallback code to helper for browser file open failure
# 04-nov-2015 guard against case mismatch in variable names

__author__  =  'spss'
__version__ =  '1.5.2'
version = __version__

import spss
import inspect, sys
import os, gettext, locale

ok1600 = spss.GetDefaultPlugInVersion()[-3:] >= '160'

# temporarily define function until Syntax class can do it right
localizationStale = True


    # debugging
    # makes debug apply only to the current thread
#try:
    #import wingdbstub
    #if wingdbstub.debugger != None:
        #import time
        #wingdbstub.debugger.StopDebug()
        #time.sleep(2)
        #wingdbstub.debugger.StartDebug()
    #import thread
    #wingdbstub.debugger.SetDebugThreads({thread.get_ident(): 1}, default_policy=0)
    # for V19 use
    #    ###SpssClient._heartBeat(False)
#except:
    #pass

try:
    _("---")
except:
    def _(msg):
        return msg

class Template(object):
    """Define a syntax element

    kwd is the keyword being defined.  It will always be made upper case, since that is how SPSS delivers keywords.
    kwd should be '' for tokens passed from a subcommand with IsArbitrary=True..
    subc is the subcommand for the keyword.  If omitted, it defaults to the (one) anonymous subcommand.
    This class does not support repeating subcommands.
    ktype is the keyword type.  Keyword type is one of
    "bool" : true,false,yes,no
    If the keyword is declared as LeadingToken in the XML and treated as bool here, the presence or absense of the keyword
      maps to True or False
    "str" : string or quoted literal with optional enumeration of valid choices.  Always converted to lower case.
    "int" : integer with optional range constraints
    "float" : float with optional range constraints
    "literal" : arbitrary string with no case conversion or validation
    "varname" : arbitrary, unvalidated variable name or syntactical equivalent (such as a dataset name)
    "existingvarlist" : list of variable names including support for TO and ALL.

    str, literal, varname, and existingvarlist are mapped to Unicode if SPSS is in Unicode mode; otherwise they are
    assumed to be code page.

    var is the Python variable name to receive the value.  If None, the lower-cased kwd value is used.  var should be
    unique across all the subcommands.  If there are duplicates, the last instance wins.  If kwd == '', var must be specified.
    vallist is the list of permitted values.  If omitted, all values of the proper type are legal.
      string values are checked in lower case.  Literals are left as written.
      For numbers, 0-2 values can be specified for any value, lower limit, upper limit.  To specify
      only an upper limit, give a lower limit of None.
    islist is True if values is a list (multiples) (SPSS keywordList or VariableNameList or NumberList)
      If islist, the var to receive the parsed values will get a list.
      existingvarlist is a list of existing variable names.  It supports SPSS TO and ALL conventions.
      """

    ktypes = ["bool", "str", "int", "float", "literal", "varname", "existingvarlist"]


    def __init__(self, kwd, subc='', var=None, ktype="str", islist=False, vallist=None):
        global _, localizationStale
        if localizationStale:
            _ = transupport(__file__,private=True)
            localizationStale = False
        if not ktype in Template.ktypes:
            localizationStale = True
            raise ValueError(_("option type must be in: ") + " ".join(Template.ktypes))
        self.ktype = ktype
        self.kwd = kwd
        self.subc = subc
        if var is None:
            self.var = kwd.lower()
        else:
            self.var = var
        self.islist = islist
        if _isseq(vallist):
            self.vallist = [u(v) for v in vallist]
        else:
            self.vallist = [u(vallist)]
        if ktype == "bool" and vallist is None:
            self.vallist = ["true", "false", "yes", "no"]
        elif ktype in  ["int", "float"]:
            if ktype == "int":
                self.vallist=[-2**31+1, 2**31-1]
            else:
                self.vallist = [-1e308, 1e308]
            try:
                if len(vallist) == 1:
                    self.vallist[0] = vallist[0]
                elif len(vallist) == 2:
                    if not vallist[0] is None:
                        self.vallist[0] = vallist[0]
                    self.vallist[1] = vallist[1]
            except:
                pass   # if vallist is None, len() will raise an exception
            
    def parse(self, item):
        key, value = item.items()[0]
        if key == 'TOKENLIST':
            key = ''   #tokenlists are anonymous, i.e., they have no keyword
        if not _isseq(value):
            value = [value]   # SPSS will have screened out invalid lists
        value = [u(v) for v in value]
        kw = self.subcdict[subc][key]  # template for this keyword
        return key, value

class ExtExistingVarlist(Template):
    """type existingvarlist"""

    def __init__(self, kwd, subc='', var=None, islist=True):
        super(ExtExistingVarlist, self).__init__(kwd=kwd, subc=subc,var=var, islist=islist)

    def parse(self, item):
        pass

class ExtBool(Template):
    "type boolean"

    def __init__(self, kwd, subc='', var=None, islist=False):
        super(ExtBool, self).__init__(kwd=kwd, subc=subc, var=var, islist=islist)
    def parse(self, item):
        key, value = super(ExtBool, self).parse(item)

def setnegativedefaults(choices, params):
    """Add explicit negatives for omitted choices if any were explicitly included.

    choices is the sequence or set of choices to consider
    params is the parameter dictionary for the command."""

    choices = set(choices)
    p = set(params)
    if p.intersection(choices):    # something was selected
        for c in choices:
            params[c] = params.get(c, False)


class Syntax(object):
    """Validate syntax according to template and build argument dictionary."""

    def __init__(self, templ, lang=None):
        """templ is a sequence of one or more Template objects.
        lang optionally specifies a language for translation.  In internal mode, lang will automatically
        match the current SPSS output language if lang is not specified here."""

        # Syntax builds a dictionary of subcommands, where each entry is a parameter dictionary for the subcommand.
        ##debugging
        #try:
            #import wingdbstub
            #if wingdbstub.debugger != None:
                #import time
                #wingdbstub.debugger.StopDebug()
                #time.sleep(2)
                #wingdbstub.debugger.StartDebug()
        #except:
            #pass

        self.unicodemode = ok1600 and spss.PyInvokeSpss.IsUTF8mode()
        if self.unicodemode:
            self.unistr = unicode
        else:
            self.unistr = str

        self.subcdict = {}
        for t in templ:
            if not t.subc in self.subcdict:
                self.subcdict[t.subc] = {}
            self.subcdict[t.subc][t.kwd] = t
        self.parsedparams = {}

        # Set up private translation for the extension module and possible translation 
        # of the parent module based on the name of the calling module.

        parent = getparent(sys._getframe(1))
        ###transupport (inspect.stack()[1][1], private=False)
        transupport(parent[0][1], private=False)
        global localizationStale
        localizationStale = True  # to force the Template class to reset the next time through


    def parsecmd(self, cmd, vardict=None):
        """Iterate over subcommands parsing each specification.

        cmd is the command specification passed to the module Run method via the EXTENSION definition.
        vardict is used if an existingvarlist type is included to expand and validate the variable names.  If not supplied,
        names are returned without validation."""

        for sc in cmd.keys():
            for p in cmd[sc]:   #cmd[sc] is a subcommand, which is a list of keywords and values
                self.parseitem(sc, p, vardict)

    def parseitem(self, subc, item, vardict=None):
        """Add parsed item to call dictionary.  

        subc is the subcommand for the item 
        item is a dictionary containing user specification.

        subc and item will already have been basically checked by the SPSS EXTENSION parser, so we can take it from there.
        If an undefined subcommand or keyword occurs (which should not happen if the xml and Template specifications are consistent), 
        a dictionary exception will be raised.
        The parsedparams dictionary is intended to be passed to the implementation as **obj.parsedparams."""

        key, value = item.items()[0]
        if key == 'TOKENLIST':
            key = ''   #tokenlists are anonymous, i.e., they have no keyword
        #value = value[0]   # value could be a list
        if not _isseq(value):
            value = [value]   # SPSS will have screened out invalid lists
        value = [u(v) for v in value]
        try:
            kw = self.subcdict[subc][key]  # template for this keyword
        except KeyError, e:
            raise KeyError(_("A syntax keyword was used that is not defined in the extension module Syntax object: %s") % e.args[0])
        if kw.ktype in ['bool', 'str']:
            value = [self.unistr(v).lower() for v in value]
            if not kw.vallist[0] is None:
                for v in value:
                    if not v in kw.vallist:
                        raise AttributeError, _("Invalid value for keyword: ") + key + ": " + v
            if kw.ktype == "str":
                self.parsedparams[kw.var] = getvalue(value, kw.islist)
            else:
                self.parsedparams[kw.var] = getvalue(value, kw.islist) in ["true", "yes", None]
        elif kw.ktype in ["varname", "literal"]:
            self.parsedparams[kw.var] = getvalue(value, kw.islist)
        elif kw.ktype in ["int", "float"]:
            if kw.ktype == "int":
                value = [int(v) for v in value]
            else:
                value = [float(v) for v in value]
            for v in value:
                if not (kw.vallist[0] <= v <= kw.vallist[1]):
                    raise ValueError, _("Value for keyword is out of range: %s") % kw.kwd
            self.parsedparams[kw.var] = getvalue(value, kw.islist)
        elif kw.ktype in ['existingvarlist']:
            self.parsedparams[kw.var] = getvarlist(value, kw.islist, vardict)
            # double check because of possible case mismatch
            varlist = self.parsedparams[kw.var]
            if not _isseq(varlist):
                varlist = [varlist]
            if vardict:
                for v in varlist:
                    if not v in vardict:
                        raise ValueError(_("Invalid variable name: %s.  Variable names are case sensitive") % v)            

def getparent(frame):
    """get the parent without resorting to entire calling stack"""
    
    # just get immediate information
    # Using 64-bit Python and 64-bit Statistics at least on Win 8, the inspect.stack api
    # sometimes fails.  It trips over an empty file name in a frame
    # This function avoids going back up the stack any more than necessary
    # Remnants of the stack.getouterframe api function remain here
    
    framelist = []
    ###while frame:  # need to limit this
    framelist.append((frame,) + inspect.getframeinfo(frame, 1))
    frame = frame.f_back
    return framelist    

def getvalue(value, islist):
    """Return value or first element.  If empty sequence, return None"""
    if islist:
        return value
    else:
        try:
            return value[0]
        except:
            return None

def getvarlist(value, islist, vardict):
    """Return a validated and expanded variable list.

    value is the tokenlist to process.
    islist is True if the keyword accepts multiples
    vardict is used to expand and validate the names.  If None, no expansion or validation occurs"""

    if not islist and len(value) > 1:
        raise ValueError, _("More than one variable specified where only one is allowed")
    if vardict is None:
        return value
    else:
        v = vardict.expand(value)
        if islist:
            return v
        else:
            return v[0]
        return 

def checkrequiredparams(implementingfunc, params, exclude=None):
    """Check that all required parameters were supplied.  Raise exception if not

    implementingfunc is the function that will be called with the output of the parse.
    params is the parsed argument specification as returned by extension.Syntax.parsecmd
    exclude is an optional list of arguments to be ignored in this check.  Typically it would include self for a class."""

    args, junk, junk, deflts = inspect.getargspec(implementingfunc)
    if not exclude is None:
        for item in exclude:
            args.remove(item)
    args = set(args[: len(args) - len(deflts)])    # the required arguments
    omitted = args - set(params)
    if omitted:
        raise ValueError(_("The following required parameters were not supplied:\n") + ", ".join(omitted))

def processcmd(oobj, args, f, excludedargs=None, lastchancef = None, vardict=None):
    """Parse arguments and execute implementation function.

    oobj is the Syntax object for the command.
    args is the Run arguments after applying
    	args = args[args.keys()[0]]
    f is the function to call to execute the command
    Whatever f returns, if anything, is returned by this function.
    excludedargs is an optional list of arguments to be ignored when checking for required arguments.
    lastchancef is an optional function that will be called just before executing the command and passed
    the parsed parameters object
    Typically it would include self for a class.
    vardict, if supplied, is passed to the parser for variable validation"""


    ##debugging
    #try:
        #import wingdbstub
        #if wingdbstub.debugger != None:
            #import time
            #wingdbstub.debugger.StopDebug()
            #time.sleep(2)
            #wingdbstub.debugger.StartDebug()
    #except:
        #pass
    
    try:
        oobj.parsecmd(args, vardict=vardict)
        # check for missing required parameters
        args, junk, junk, deflts = inspect.getargspec(f)
        if deflts is None:   #getargspec definition seems pretty dumb here
            deflts = tuple()
        if not excludedargs is None:
            for item in excludedargs:
                args.remove(item)
        args = set(args[: len(args) - len(deflts)])    # the required arguments
        omitted = [item for item in args if not item in oobj.parsedparams]
        if omitted:
            raise ValueError, _("The following required parameters were not supplied:\n") + ", ".join(omitted)
        if not lastchancef is None:
            lastchancef(oobj.parsedparams)
        return f(**oobj.parsedparams)
    except:
        # Exception messages are printed here as a pivot table, 
        # but the exception is not propagated, and tracebacks are suppressed,
        # because as an Extension command, the Python handling should be suppressed.
        
        # But, if the SPSS_EXTENSIONS_RAISE environment variable exists and has the value "true"
        # the exception is reraised instead.

        ###raise   #debug
        if 'SPSS_EXTENSIONS_RAISE' in os.environ and os.environ['SPSS_EXTENSIONS_RAISE'].lower() == "true":
            raise
        else:
            myenc = locale.getlocale()[1]  # get current encoding in case conversions needed        
            warnings = NonProcPivotTable("Warnings",tabletitle=_("Warnings "))
            msg = sys.exc_info()[1]
            if _isseq(msg):
                # try to make error message into something a pivot table will understand
                # avoid forcing a str conversion if possible, but numbers and classes need
                # to be converted
                #msg = ",".join([(isinstance(item, (float, int)) or item is None) and str(item) or \
                    #(isinstance(item, Exception) and str(item)) or item for item in msg])
                msg = ",".join([unicodeit(item, myenc) for item in msg])
            if len(msg) == 0:   # no message with exception
                msg = str(sys.exc_info()[0])  # if no message, use the type of the exception (ugly)
            warnings.addrow(msg)
            sys.exc_clear()
            warnings.generate()

def unicodeit(value, myenc):
    if isinstance(value, (int, float)):
        return unicode(value)
    if value is None:
        return ""
    if not isinstance(value, unicode):  # could be str or a class
        try:
            return  unicode(value, myenc)
        except:
            return unicode(str(value), myenc)
    else:
        return value
        
def floatex(value, format=None):
    """Return value as a float if possible after addressing format issues

    value is a (unicode) string that may have a locale decimal and other formatting decorations.
    raise exception if value cannot be converted.
    format is an optional format specification such as "#.#".  It is used to disambiguate values
    such as 1,234.  That could be either 1.234 or 1234 depending on whether comma is a
    decimal or a grouping symbol.  Without the format, it will be treated as the former.
    This function cannot handle date formats.  Such strings will cause an exception.
    A sysmis value "." will cause an exception."""

    try:
        return float(value)
    except:
        if format == "#.#":
            #  comma must be the decimal and  no other decorations may be present
            value = value.replace(",", ".")
            return float(value)
        # maybe a comma decimal or COMMA format
        lastdot = value.rfind(".")
        lastcomma = value.rfind(",")
        if lastcomma > lastdot:  # handles DOT format and F or E with comma decimal
            value = value.replace(".", "")
            value = value.replace(",", ".")
        elif lastdot > lastcomma:  # truly a dot decimal format
            value = value.replace(",", "")   # handles COMMA format	    
        v = value.replace(",", ".")
        try:
            return float(v)
        except:
            # this is getting annoying.  Maybe a decorated format.  "/" is included below
            # to ensure that conversion will fail for date formats

            v = "".join([c for c in value if c.isdigit() or c in ["-", ".", "+", "e","E", "/"]])
            return float(v)   # give up if this fails

# The following routines are copied from spssaux in order to avoid the need to import that entire module
def u(txt):
    """Return txt as Unicode or unmodified according to the SPSS mode"""

    if not ok1600 or not isinstance(txt, str):
        return txt
    if spss.PyInvokeSpss.IsUTF8mode():
        if isinstance(txt, unicode):
            return txt
        else:
            return unicode(txt, "utf-8")
    else:
        return txt

def _isseq(obj):
    """Return True if obj is a sequence, i.e., is iterable.

    Will be False if obj is a string, Unicode string, or basic data type"""

    # differs from operator.isSequenceType() in being False for a string

    if isinstance(obj, basestring):
        return False
    else:
        try:
            iter(obj)
        except:
            return False
        return True

def transupport(thefile,localedir=None, lang=None, private=False):
    """Return a function that retrieves translated strings and optionally install it in builtins.

    thefile is the full path of the module.  By default, translations are expected to live 
    below this location in a /lang subdirectory with further structure
    /de or other posix language/LC_MESSAGES/module-name.mo.
    If localedir is specified, the language-specific directories are assumed
    to be under that location starting with the LC_MESSAGES directory.
    
    Example.
    If the implementing module is installed in c:/spss18/extensions and localedir is not specified,
    c:/spss18/extensions/SPSSINC_CENSOR_TABLES/lang/de/LC_MESSAGES/SPSSINC_CENSOR_TABLES.mo
    would hold the German (de) translation table for the SPSSINC_CENSOR_TABLES extension command.
    
    lang can specify a specific language code for translations.  If not specified and
    in internal mode, it will be synchronized with the current SPSS output language for version
    18 and later.  In external mode, the gettext-relevant environment variables can 
    control the language or it can be passed explicitly, or, better, 
    submit a SET OLANG command first.
    If the language files are not found, messages are left untranslated.
    If private, the translation function is returned.  Otherwise it is installed in 
    builtins and can be used everywhere.
    
    For V17, you can set the LANGUAGE environment variable at the os level in order to control
    the output language for extension commands statically.
    
    In order to avoid conflict with the _ behavior of the interpreter, assignment of interactive
    values to _ is suppressed.
    """
    
    # SPSS settings for simplified and traditional chinese are not recognized by Python.
    # Furthermore, the Python gettext library looks at the LANGUAGE environment variable.
    # So here we change that variable to match the directory structure we use for these
    # languages.  This change should not be seen by the parent process, but the setting is
    # restored in case external mode is being used.
    
    langcodes = {"schinese":"zh_CN", "tchinese":"zh_TW", "bportugu": "pt_BR"}
    origlang = os.environ.get("LANGUAGE")
    if lang is None:
        try:
            lang=os.environ["LANGUAGE"].lower()
        except:
            lang="english"

    if lang in langcodes:
        os.environ["LANGUAGE"] = langcodes[lang]
    lang = langcodes.get(lang, lang)
    thename = os.path.basename(os.path.splitext(thefile)[0])
    sys.displayhook = _dh  #suppress _ assignment
    if localedir is None:
        localedir=os.path.dirname(thefile) + "/" + thename  + "/lang"
    if private:
        tr = gettext.translation(thename, fallback=True, localedir=localedir,
            languages=[lang])
        if origlang:
            os.environ["LANGUAGE"] = origlang
        return tr.ugettext
    else:
        gettext.install(thename, localedir=localedir, unicode=True)
        if origlang:
            os.environ["LANGUAGE"] = origlang

def _dh(obj):
    """Function to override interactive expression display, disabling assignment to _"""
    if not obj is None:
        print repr(obj)

class NonProcPivotTable(object):
    """Accumulate an object that can be turned into a basic pivot table once a procedure state can be established"""
    
    def __init__(self, omssubtype, outlinetitle="", tabletitle="", caption="", rowdim="", coldim="", columnlabels=[],
                 procname="Messages"):
        """omssubtype is the OMS table subtype.
        caption is the table caption.
        tabletitle is the table title.
        columnlabels is a sequence of column labels.
        If columnlabels is empty, this is treated as a one-column table, and the rowlabels 
        are used as the values with the label column hidden
        
        procname is the procedure name.  It must not be translated."""
        
        attributesFromDict(locals())
        self.rowlabels = []
        self.columnvalues = []
        self.rowcount = 0

    def addrow(self, rowlabel=None, cvalues=None):
        """Append a row labelled rowlabel to the table and set value(s) from cvalues.
        
        rowlabel is a label for the stub.
        cvalues is a sequence of values with the same number of values are there are columns in the table."""
        
        if cvalues is None:
            cvalues = []
        self.rowcount += 1
        if rowlabel is None:
            self.rowlabels.append(str(self.rowcount))
        else:
            self.rowlabels.append(rowlabel)
        self.columnvalues.extend(cvalues)
        
    def generate(self):
        """Produce the table if it has any rows, assuming that a procedure state is now in effect or possible"""
        
        privateproc = False
        if self.rowcount > 0:
            try:
                table = spss.BasePivotTable(self.tabletitle, self.omssubtype)
            except:
                spss.EndDataStep()  # just in case there is a dangling DataStep
                spss.StartProcedure(self.procname)
                privateproc = True
                table = spss.BasePivotTable(self.tabletitle, self.omssubtype)
            if self.caption:
                table.Caption(self.caption)
            if self.columnlabels != []:
                table.SimplePivotTable(self.rowdim, self.rowlabels, self.coldim, 
                    self.columnlabels, self.columnvalues)
            else:
                table.Append(spss.Dimension.Place.row,"rowdim",hideName=True,hideLabels=True)
                table.Append(spss.Dimension.Place.column,"coldim",hideName=True,hideLabels=True)
                colcat = spss.CellText.String("Message")
                for r in self.rowlabels:
                    if isinstance(r, (int, float)):
                        r = str(r)
                    cellr = spss.CellText.String(r)
                    table[(cellr, colcat)] = cellr
            if privateproc:
                spss.EndProcedure()
                
def attributesFromDict(d):
    """build self attributes from a dictionary d."""
    self = d.pop('self')
    for name, value in d.iteritems():
        setattr(self, name, value)
def _isseq(obj):
    """Return True if obj is a sequence, i.e., is iterable.
    
    Will be False if obj is a string or basic data type"""
    
    # differs from operator.isSequenceType() in being False for a string
    
    if isinstance(obj, basestring):
        return False
    else:
        try:
            iter(obj)
        except:
            return False
        return True
    
def helper(helpfile="markdown.html"):
    """open html help in default browser window
    
    The location is computed from the current module name
    This function must be called only from the extension command's Run function
    The default help file is expected to be named markdown.html and to
    be in the directory matching the command name.
    
    helpfile can specify an alternate name for the help file"""
    
    import webbrowser, os.path

    path = os.path.splitext(getparent(sys._getframe(1))[0][1])[0]
    helpspecbase =  path + os.path.sep + helpfile
    helpspec = "file:///" + helpspecbase
    #helpspec = "file://" + path + os.path.sep + \
         #helpfile
    
    # webbrowser.open seems not to work well
    try:
        browser = webbrowser.get()
    except:
        raise SystemError("""No runnable browser was found to display help.  Open the help file manually
in an appropriate program: %s""" % helpspec)
    if not browser.open_new(helpspec):
        try:
            os.startfile(helpspecbase)   # Windows only
        except:
            print("Help file not found:" + helpspecbase)

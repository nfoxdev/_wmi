*-----------------------------------------------------------------------------------
* Marco Plaza, 2022
* https://github.com/nfoxdev
* this program should be saved as _WMI.PRG
*-----------------------------------------------------------------------------------
*
* simple usage: 
* _wmi( wmiClass [where <filter condition>] [, wmiNameSpace] )
* where: optional query filter
* wmiNameSpace defaults to "CIMV2"
*
* ie: 
*
* oDisks = _wmi('Win32_diskDrive') 
* oMonitors = _wmi('Win32_PNPEntity where service = "monitor"')
*
* Test: save this program as _wmi.prg and do "testme in _wmi"
*
*-----------------------------------------------------------------------------------------
lparameters wmiclass as String,wminamespace as string,includenulls as Boolean as Object


local oerror,wmiclass,wminamespace,emessage,objwmiservice,oquery,owmi

wmiclass     = evl(m.wmiclass,'')
wminamespace = evl(m.wminamespace,'CIMV2')
emessage     = ''


try
   objwmiservice = getobject("winmgmts://./root/"+m.wminamespace)
   oquery   = objwmiservice.execquery( 'SELECT * FROM '+m.wmiclass,,48)
   owmi   = processobject( m.oquery , m.includenulls )
catch to oerror
   owmi = null
endtry

if _error(m.oerror)
   error ' Invalid WMI Class or NameSpace '
endif

return m.owmi

*-------------------------------------------------
procedure processobject( oquery , includenulls )
*-------------------------------------------------
local oerror,owmi,nitem,oitem

owmi = _newempty('items(1)')

nitem = 0

try

   for each oitem in m.oquery

      nitem = m.nitem + 1
      dimension owmi.items(m.nitem)
      owmi.items(m.nitem) = createobject('empty')
      setproperties( m.oitem, owmi.items(m.nitem), m.includenulls )

   endfor

   addproperty(owmi,'count',m.nitem)

catch to oerror

   owmi = null

endtry

if vartype(m.oerror) = 'O'
   error m.oerror.message
endif


return m.owmi

*--------------------------------------------------------
procedure setproperties( oitem , otarget , includenulls )
*--------------------------------------------------------

local oerr,thisproperty,thisarray,nitem,thisitem,newname,property,item

for each property in m.oitem.properties_

   do case
   case isnull( m.property.value) and !m.includenulls
      loop

   case vartype( m.property.value ) = 'O'
      thisproperty = createobject('empty')
      setproperties(m.property.value, m.thisproperty )
      addproperty( otarget ,m.property.name,m.thisproperty)

   case m.property.isarray
      addproperty( otarget ,property.name+'(1)',.null.)
      thisarray = 'otarget.'+m.property.name

      nitem = 0

      if !isnull(m.property.value)

         for each item in m.property.value

            nitem = m.nitem+1
            dimension &thisarray(m.nitem)

            if vartype( m.item) = 'O'
               thisitem = createobject('empty')
               setproperties( m.item, m.thisitem )
               &thisarray(m.nitem) = m.thisitem
            else
               &thisarray(m.nitem) = m.item
            endif

         endfor

      endif

   otherwise
      try
         addproperty( otarget ,m.property.name,m.property.value)
      catch
         newname =  _name2validname(property.name)
         addproperty( otarget ,m.newname,m.property.value)
      endtry

   endcase

endfor


*----------------------------------
procedure testme
*----------------------------------
public oinfo

oinfo = create('empty')

wait 'Running WMI Query....please wait.. ' window nowait at wrows()/2,wcols()/2

addproperty( oinfo, "monitors"  , _wmi('Win32_PNPEntity where service = "monitor"') )
addproperty( oinfo, "diskdrive" , _wmi('Win32_diskDrive') )
addproperty( oinfo, "startup" ,   _wmi('Win32_startupCommand'))
addproperty( oinfo, "BaseBoard" , _wmi('Win32_baseBoard') )
addproperty( oinfo, "netAdaptersConfig",  _wmi('Win32_NetworkAdapterConfiguration') )


messagebox( 'Use Intellisense/debugger to check oInfo',0)
debug


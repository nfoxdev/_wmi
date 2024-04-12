*-----------------------------------------------------------------------------------
* Marco Plaza, 2022
* https://github.com/nfoxdev/_wmi
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
*------------------------------------------------------------------------------------
lparameters wmiquery,wmiclass

local oerr,emessage,objwmiservice,oquery,owmi

wmiclass = Evl(m.wmiclass,'CIMV2')
wmiquery = Evl(m.wmiquery,'')

emessage = ''

Try
   objwmiservice = Getobject("winmgmts:\\.\root\"+m.wmiclass)
   oquery = objwmiservice.execquery( 'SELECT * FROM '+m.wmiquery,,48)
   owmi = processobject( oquery )
Catch To oerr
   emessage = m.oerr.Message
Endtry

If !Empty(m.emessage)
   Error ' Invalid Query/WmiClass '
   Return .Null.
Else
   Return m.owmi
Endif

*-------------------------------------------------
Procedure processobject( oquery )
*-------------------------------------------------
local owmi,nitem,oitem

owmi = Createobject('empty')
AddProperty(owmi,'items(1)',.Null.)
nitem = 0

Try

   For Each oitem In m.oquery

      nitem = m.nitem + 1
      Dimension owmi.items(m.nitem)
      owmi.items(m.nitem) = Createobject('empty')
      setproperties( m.oitem, owmi.items(m.nitem) )

   Endfor

Catch

Endtry

AddProperty(owmi,'count',m.nitem)

Return m.owmi

*--------------------------------------------------------
Procedure setproperties( oitem , otarget  )
*--------------------------------------------------------

local oerr,thisproperty,thisarray,nitem,thisitem,property,item

For Each property In m.oitem.properties_
   Try
      Do Case
      Case Vartype( m.property.Value ) = 'O'
         thisproperty = Createobject('empty')
         setproperties(m.property.Value, m.thisproperty )
         AddProperty( otarget ,m.property.Name,m.thisproperty)

      Case m.property.isarray

         AddProperty( otarget ,property.Name+'(1)',.Null.)
         thisarray = 'otarget.'+m.property.Name

         nitem = 0

         If !Isnull(m.property.Value)

            For Each Item In m.property.Value

               nitem = m.nitem+1
               Dimension &thisarray(m.nitem)

               If Vartype( m.item) = 'O'
                  thisitem = Createobject('empty')
                  setproperties( m.item, m.thisitem )
                  &thisarray(m.nitem) = m.thisitem
               Else
                  &thisarray(m.nitem) = m.item
               Endif

            Endfor

         Endif

      Otherwise
         AddProperty( otarget ,m.property.Name,m.property.Value)
      Endcase

   Catch To oerr
      Messagebox(Message(),0)
   Endtry
Endfor


*----------------------------------
procedure testme
*----------------------------------
Public oinfo

oinfo = Create('empty')

Wait 'Running WMI Query....please wait.. ' Window Nowait At Wrows()/2,Wcols()/2


addproperty( oinfo, "monitors"  , wmiquery('Win32_PNPEntity where service = "monitor"') )
addproperty( oInfo, "diskdrive" , wmiquery('Win32_diskDrive') )
addproperty( oInfo, "startup" ,   wmiquery('Win32_startupCommand'))
addproperty( oInfo, "BaseBoard" , wmiquery('Win32_baseBoard') )
addproperty( oInfo, "netAdaptersConfig",  wmiquery('Win32_NetworkAdapterConfiguration') )


Messagebox( 'Please explore "oInfo" in debugger watch window or command line ',0)



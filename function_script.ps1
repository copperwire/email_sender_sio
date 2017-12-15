
Function sendMail($addr_to, $attach)
{	  

	$olMailItem = 0  
	  
	$olApp = new-object -comobject outlook.application  
	  
	$NewMail = $olApp.CreateItem($olMailItem)  
	  
	$NewMail.SentonBehalfofName = "foreninger@sio.no"
	
	$NewMail.Subject = "Liste over interesserte etter forenigsdagen"  
	  
	$NewMail.To = $addr_to

	$NewMail.Attachments.Add($attach)

	$NewMail.Body = "Hei, `n Vedlagt finner  du en excel fil med alle studentene som viste interesse for din forening under foreningsdagen. Hvis de interesserte ikke har oppgitt telefonnummer vil det fremkomme som "!NUM" i excel-arket. `n Med vennlig hilsen `n SiO Foreninger"

	$NewMail.Send()

	$olApp.Quit()

	return 0
}

Function sendMail_noleads($addr_to)
{	  

	$olMailItem = 0  
	  
	$olApp = new-object -comobject outlook.application  
	  
	$NewMail = $olApp.CreateItem($olMailItem)  

	$NewMail.SentonBehalfofName = "foreninger@sio.no"
	  
	$NewMail.Subject = "Foreningsdagen"  
	  
	$NewMail.To = $addr_to

	$NewMail.Body = "Hei, `n Under Foreningsdagen samlet vi inn foreningsønsker fra de besøkende gjennom et nettskjema. Vi merket dog at de fleste tok kontakt direkte med dere foreninger. Dere fikk ingen supplerende interessenter etter foreningsdagen. `n Med vennlig hilsen `n SiO Foreninger"

	$NewMail.Send()

	$olApp.Quit()

	return 0
}
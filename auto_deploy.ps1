#####################################################################################################################
#   ESXi 5.5 Auto Deploy VM Clone                                                               Create 2017/11/30   #
#   Author Annda Ver 1.0.0                                                                      Update 2017/12/01   #
#####################################################################################################################
# Configure																											#
#####################################################################################################################
$vcenterhost = "192.168.1.254"
$vcenterusr = "root"
$vcenterpsk = "password"
$spec="auto-deploy"
Stop-Process -Name excel
#####################################################################################################################
# vCenter Connect																									#
#####################################################################################################################
Connect-VIServer -Server $vcenterhost -username $vcenterusr -Password $vcenterpsk
$custsysprep = Get-OSCustomizationSpec $spec
#####################################################################################################################
# Read Excel																										#
#####################################################################################################################
# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application
# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $false
#Specify the path of the excel file
$FilePath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$FilePath = $FilePath + "\host_list.xlsx"
# Open the Excel file(ReadOnly mode) and save it in $WorkBook  
$WorkBook = $objExcel.Workbooks.Open($FilePath, $true)  
# Load the WorkSheet  
$WorkSheet = $WorkBook.Sheets.Item(1)
#####################################################################################################################
# Clone VM																											#
#####################################################################################################################
$esxihost = "auto deploy"; $x = 1
do
{
    $x = $x + 1
    $esxihost = $WorkSheet.cells.item($x,1).Text -replace "\s{1,}",""
    $vmname = $WorkSheet.cells.item($x,2).Text -replace "\s{1,}",""
    $template = $WorkSheet.cells.item($x,3).Text -replace "\s{1,}",""
    $datastore = $WorkSheet.cells.item($x,4).Text -replace "\s{1,}",""
    $hostname = $WorkSheet.cells.item($x,5).Text -replace "\s{1,}",""
    $ip = $WorkSheet.cells.item($x,6).Text -replace "\s{1,}",""
    $mask = $WorkSheet.cells.item($x,7).Text -replace "\s{1,}",""
    $gateway = $WorkSheet.cells.item($x,8).Text -replace "\s{1,}",""
    $net_tag = $WorkSheet.cells.item($x,9).Text -replace "\s{1,}",""
    if ( $esxihost -ne "" )
    {
        $custsysprep | Set-OScustomizationSpec -NamingScheme fixed -NamingPrefix $hostname
        $custsysprep | Get-OSCustomizationNicMapping | Set-OSCustomizationNicMapping -IpMode UseStaticIP -IpAddress $ip -SubnetMask $mask -DefaultGateway $gateway
        New-vm -vmhost $esxihost -Name $vmname -Template $template -Datastore $datastore -OSCustomizationspec $custsysprep
        if ( $? -eq "true" )
        {
            Get-VM $vmname | Get-NetworkAdapter | Set-NetworkAdapter -NetworkName $net_tag -Confirm:$false
            if ( $? -eq "true" )
            {
                $WorkSheet.cells.item($x,10) = "Success"
                start-vm -vm $vmname
            }
            else
            {
                $WorkSheet.cells.item($x,10) = "Failed"
            }
        }
        else
        {
            $WorkSheet.cells.item($x,10) = "Failed"
        }
	$WorkBook.Save() | Out-Null
    }
}
while($esxihost -ne "")
#####################################################################################################################
# Scripts Exit																										#
#####################################################################################################################
$WorkBook.Close()  
$objExcel.Quit()
Stop-Process -Name excel
"ESXi 5.5 Auto Deploy Clone VM Complete."
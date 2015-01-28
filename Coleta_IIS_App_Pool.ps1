<#	 
	============================================================================================================ 
	.Nome SCRIPT: Coleta_IIS_App_Pool.ps1 
    .DESCRIÇÃO:Script Coletar informacoes do application pools de servidores IIS. 
    .PARAMETROS: <Nenhum> 
    .Entradas(INPUTs):
     - Caminho do arquivo de texto contendo listagem de computadores a serem verficados
     
     .Saida (OUTPUTs): 
      - Arquivo CSV gravado na pasta onde o script esta localizado
       
     .Notas: 
      Versao: 1.0 
      Criado em: 28/01/2015 
      Autores:
              Renato Pagan
              Warley Penido Ferreira 
      Email:  warley.penido@gmail.com 
     ==========================================================================================================
 #> 


#------------------------------------------[Entradas <INPUTS>]-------------------------------------------------


$Computers = Get-Content -Path C:\Warley\Computers.txt


#--------------------------------------[Inicialisacoes / Declaracoes]------------------------------------------

$local_script_path = $(Split-Path $MyInvocation.InvocationName)
$script_name_without_extension = [io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)
$output_file = $local_script_path +"\"+ $script_name_without_extension +"_"+ (Get-Date -UFormat "%Y%m%d_%Hh%M") +".csv"
$output_arr = @()
$msg = $null


#-----------------------------------------------[Execucao]-----------------------------------------------------

foreach ($ComputerName in $Computers){
    try {
        if ( (gwmi -query "select name from __namespace where name = 'MicrosoftIISv2'" -Namespace root -ComputerName $ComputerName) -ne "" ) { 

            try {
                $sites = gwmi -ComputerName $ComputerName -Namespace "root\MicrosoftIISv2" -Class "IISWebServerSetting" -Authentication 6 | select ServerComment, Name, ServerBindings
                $apps = gwmi -ComputerName $ComputerName -Namespace "root\MicrosoftIISv2" -Class "IIsApplicationPoolSetting" -Authentication 6

                foreach ($site in $sites) {
                    foreach ( $app in $apps ) {
                        if ( $app.Name -ne $Null) {
                            $obj = New-Object PSObject -Property @{
                                    ComputerName = $ComputerName
                                    Site         = $site.ServerComment
                                    App          = $app.Name
                                    DotNet       = $runtime
                                    UserName     = $app.WAMUserName
                                    Status       = "OK"
                                   }
                            $output_arr += $obj
                        }
                    }
                }
            } 
            catch {
                $msg = "$ComputerName::ERRO: "+ $error[0].Exception.Message
                $obj_ne = New-Object PSObject -Property @{
                        ComputerName = $ComputerName
                        Status = $msg
                }
                $output_arr += $obj_ne
            }

        } 
        else {
            $msg = "$ComputerName::ERRO: "+ $error[0].Exception.Message
            $obj_ne = New-Object PSObject -Property @{
                    ComputerName = $ComputerName
                    Status = $msg
            }
            $output_arr += $obj_ne
        }

    }
    catch{    
            $msg = "$ComputerName::ERRO: "+ $error[0].Exception.Message
            $obj_ne = New-Object PSObject -Property @{
                    ComputerName = $ComputerName
                    Status = $msg
            }
            $output_arr += $obj_ne
    }
    $msg
    $output_arr | select ComputerName, Site, App, UserName, Status | Export-Csv -Path ($output_file) -noTypeInformation
}

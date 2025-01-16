import pyautogui
import win32com.client
import subprocess
import time



def conectar_sap():
    try:
        # Inicializa o SAP GUI
        SapGuiAuto = win32com.client.GetObject("SAPGUISERVER")
        if not SapGuiAuto:
            raise Exception("SAP GUI não encontrado. Certifique-se de que está aberto.")

        application = SapGuiAuto.GetScriptingEngine

        # Conecta à primeira conexão disponível (ou especifique o nome da conexão)
        connection = application.Children(0)

        # Seleciona a primeira sessão
        session = connection.Children(0)

        # Interações no SAP GUI
        session.findById("wnd[0]").resizeWorkingPane(235, 50, False)
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "100"  # Cliente
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "0931221"  # Usuário
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "@Floki1046"  # Senha
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"  # Idioma
        session.findById("wnd[0]/usr/txtRSYST-LANGU").setFocus()
        session.findById("wnd[0]/usr/txtRSYST-LANGU").caretPosition = 2
        session.findById("wnd[0]").sendVKey(0)

        print("Login realizado com sucesso!")
        return session

    except Exception as e:
        print(f"Erro ao conectar ao SAP GUI: {e}")
        return None

def iniciar_sap():
    try:
        subprocess.Popen([r'C:\Program Files (x86)\SAP\NWBC65\NWBC'])
        print("Abrindo o SAP")

        imagem = 'gatilho_login.png'

        while True:
            try:
                pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9)
                print('SAP carregado. Continuando...')
                break
            except:
                time.sleep(1)
                print('Procurando SAP...')

        print('SAP iniciado com sucesso. Pronto para logar.')
    except Exception as e:
        print(f'Erro ao iniciar o SAP: {e}')

def script_me5a(session):
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "ME5A"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = "ENUC_GERAL"
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
        session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
        session.findById("wnd[1]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 109
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem("&FILTER")
        session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "/ENUC_RPA_01"
        session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 12
        session.findById("wnd[2]").sendVKey(0)
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\tesch\Documents\teste_extracao_rpa"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME5A.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        print('ME5A executada com sucesso! \n  > Passando para a próxima transação.')
    except:
        print('Falha ao tentar executar a transação ME5A ')



def mm60(session):
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "mm60"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()


        try:
            session.findById("wnd[0]/tbar[1]/btn[33]").press()

        except Exception as e:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            print(f'Transação não disponível...{e}')


        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setColumnWidth("VARIANT", 19)
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem("&FILTER")
        session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "/ENUC_CATALO"
        session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 12
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\tesch\Documents\teste_extracao_rpa"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MM60.txt"
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
        session.findById("wnd[0]").sendVKey(0)

        print('MM60 executada com sucesso! \n\n  > Fim do programa.')
    except:
        print('Falha ao tentar executar a transação MM60')
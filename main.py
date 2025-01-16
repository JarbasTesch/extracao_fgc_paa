import funcoes as fcs

fcs.iniciar_sap()
session = fcs.conectar_sap()
fcs.script_me5a(session)
fcs.mm60(session)
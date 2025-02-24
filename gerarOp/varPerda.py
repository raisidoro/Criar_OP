

def iniPerdas(setor):
    classificaDefeitos = ""
    if(setor == "IE-INJECAO"):
        classificaDefeitos = {
            "RI":0,
            "RE":0,
            "MT":0,
            "CO":0,
            "PNDEF":0,
            "FI":0,
            "CG":0,
            "MG":0,
            "RB":0,
            "PNALOJMD":0,
            "STP":0
        }
    elif(setor == "IE-MONTAGEM"):
        classificaDefeitos = {
            "CD":0,
            "PE":0,
            "SOLDBX":0,
            "RP":0,
            "RI":0,
            "ST":0,
            "SP":0,
            "VS":0,
            "DEFDEINJ":0,
            "DEFDPIN":0,
            "STP":0
        }
    elif(setor == "IE-PINTURA"):
        classificaDefeitos = {
            "FO":0,
            "SJ":0,
            "ET":0,
            "RI":0,
            "RG":0,
            "DEFDEINJ":0,
            "MC":0,
            "PB":0,
            "CT":0,
            "MA":0,
            "STP":0
        }
    elif(setor == "SS-MONTAGEM"):
        classificaDefeitos = {
           "CMPNTNG":0,
            "FLPRC":0,
            "GRRBOLNG":0,
            "QDDECMP":0,
            "QE":0,
            "QDDEPROD":0,
            "QS":0,
            "REIMETQ":0,
            "TSTELENG":0,
            "OT":0,
            "STP":0
        }
    elif(setor == "WS-CLIP"):
        classificaDefeitos = {
            "ESCTRMQ":0,
            "EC":0,
            "FLHAFURO":0,
            "FHSENEME":0,
            "FALTCLIP":0,
            "FALTCOLA":0,
            "PFRASG":0,
            "QE":0,
            "ANORANT":0,
            "OT":0,
            "STP":0
        }
    elif(setor == "WS-DW"):
        classificaDefeitos = {
           "FB":0,
            "RL":0,
            "ER":0,
            "FL":0,
            "PM":0,
            "PCDOBR":0,
            "RESCOLA":0,
            "RBNMOLDE":0,
            "PU":0,
            "ANORANT":0,
            "STP":0
        }
    elif(setor == "WS-MIX"):
        classificaDefeitos = {
            "EI":0,
            "SemClassif1":0,
            "SemClassif2":0,
            "SemClassif3":0,
            "SemClassif4":0,
            "SemClassif5":0,
            "SemClassif6":0,
            "SemClassif7":0,
            "SemClassif8":0,
            "SemClassif9":0,
            "STP":0
        }
    # elif(setor == "WS-SPONGE"):
    #     classificaDefeitos = {
    #         "FM":0,
    #         "FC":0,
    #         "IQ":0,
    #         "RL":0,
    #         "PI":0,
    #         "PFLFORA":0,
    #         "ENTBCPIN":0,
    #         "FRORGA":0,
    #         "RLTCLIP":0,
    #         "PCDRAN":0,
    #         "STP":0
    #     }
    elif(setor == "WS-SPONGE"):
        classificaDefeitos = {
            "PRDINCMN":0,
            "PFNG":0,
            "SRRLHADO":0,
            "FLHAPTIN":0,
            "FURONG":0,
            "CORTENG":0,
            "BOLHAS":0,
            "FLHALASE":0,
            "IM":0,
            "":0,
            "STP":0
        }
    elif(setor == "WS-TPV"):
        classificaDefeitos = {
           "IM":0,
            "DF":0,
            "RM":0,
            "ON":0,
            "AF":0,
            "FLHEQUIP":0,
            "PC":0,
            "SRRLHDO":0,
            "BOLHA":0,
            "OT":0,
            "STP":0
        }
    elif(setor == "WS-GR"):
        classificaDefeitos = {
            "DS":0,
            "PD":0,
            "FL":0,
            "PM":0,
            "TP":0,
            "ER":0,
            "ERROPE":0,
            "PRBLTAPE":0,
            "ANORANT":0,
            "OT":0,
            "STP":0
        }
    elif(setor == "WS-OT"):
        classificaDefeitos = {
            "FM":0,
            "FL":0,
            "ER":0,
            "AM":0,
            "MD":0,
            "PG":0,
            "CTNOBULB":0,
            "FC":0,
            "DS":0,
            "ANORANT":0,
            "STP":0
        }

    return classificaDefeitos




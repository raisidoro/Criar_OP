

def iniRetrabalho(setor):
    classificaRetrabalho = ""
    if(setor == "IE-INJECAO"):
        classificaRetrabalho = {
            "R-CG":0
        }
    elif(setor == "IE-MONTAGEM"):
        classificaRetrabalho = {
            "R-CD":0,
            "R-PE":0,
            "R-SOLDBX":0,
            "R-RI":0,
            "R-ST":0,
            "R-VS":0,
            "R-DFDINJ":0,
            "R-DFDPIN":0
        }
    elif(setor == "IE-PINTURA"):
        classificaRetrabalho = {
            "R-FO":0,
            "R-SJ":0,
            "R-ET":0,
            "R-RI":0,
            "R-RG":0,
            "R-DFDINJ":0,
            "R-MC":0,
            "R-CT":0,
            "R-MA":0            
        }
    elif(setor == "SS-MONTAGEM"):
        classificaRetrabalho = {
            "R-RIMETQ":0,
            
        }
    elif(setor == "WS-CLIP"):
        classificaRetrabalho = {
            "R-EC":0,
            "R-FLAFRO":0,
            "R-FLCLIP":0,
            "R-FLTCLA":0,
            "R-PFRASG":0            
        }
    elif(setor == "WS-DW"):
        classificaRetrabalho = {
            "R-FB":0,
            "R-RL":0,
            "R-ER":0,
            "R-FL":0,
            "R-PM":0,
            "R-PCDOBR":0,
            "R-RESCOL":0,
            "R-RBNMLD":0,
            "R-PU":0           
        }
    elif(setor == "WS-MIX"):
        classificaRetrabalho = {
            "SemClassif1":0
        }
    elif(setor == "WS-SPONGE"):
        classificaRetrabalho = {
            "R-FC":0,
            "R-PCDRAN":0
        }
    elif(setor == "WS-TPV"):
        classificaRetrabalho = {
            "SemClassif1":0,
            
        }
    elif(setor == "WS-GR"):
        classificaRetrabalho = {
            "SemClassif1":0,
           
        }
    elif(setor == "WS-OT"):
        classificaRetrabalho = {
            "R-FM":0,
            "R-FL":0,
            "R-ER":0,
            "R-AM":0,
            "R-MD":0,
            "R-PG":0,
            "R-CTNBLB":0,
            "R-FC":0,
            "R-DS":0,
        }

    return classificaRetrabalho




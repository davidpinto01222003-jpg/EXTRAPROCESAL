"""Estimación alternativa para retoques livianos: una frase levemente editada hereda
la marca real de la frase del último informe de la que viene (ver CAMBIOS_COMPILATIO.md)."""
import sys,pickle,numpy as np,collections; sys.path.insert(0,'/home/user/EXTRAPROCESAL/simulador_compilatio'); import simulador as S
def lab(d,p):
    pars=S.leer_docx(d); pc=S.cuerpo(pars); ia,_,_=S.marcas_pdf(p); pares,y=S.etiquetar(pc,ia)
    return [(s,int(v)) for (_,s,_),v in zip(pares,y)], pars
def heredar(s, viejas):
    """marca heredada: promedio ponderado de las frases viejas cuyo texto quedó contenido en s"""
    sh=S.shingles(s); 
    if not sh: return None, 0
    tot=0; acc=0; cub=set()
    for t,v in viejas:
        st=S.shingles(t)
        if not st: continue
        inter=st & sh
        if len(inter)/len(st)>=0.5:
            acc+=v*len(inter); tot+=len(inter); cub|=inter
    if not tot: return None, 0
    return acc/tot, len(cub)/len(sh)
if __name__=="__main__":
    # validación: v4 -> v6 (ambos informes completos)
    L4,_=lab(S.AQUI/"datos/v4.docx",S.AQUI/"datos/reporte_compilatio_v4.pdf"); L6,_=lab(S.AQUI/"datos/v6.docx",S.AQUI/"datos/reporte_compilatio_v6.pdf")
    s4={s for s,_ in L4}; res=[]
    for s,v in L6:
        if s in s4: continue
        h,c=heredar(s,L4)
        if h is not None and c>=0.6: res.append((round(h),v,len(s.split())))
    c=collections.Counter((a,b) for a,b,_ in res); n=len(res)
    print(f"v4→v6, frases retocadas (≥60% del texto viene de la v4): {n}; la marca se conserva en {(c[(0,0)]+c[(1,1)])/n:.0%}  {dict(c)}")
    w=lambda f: sum(x for a,b,x in res if f(a,b))
    print(f"  palabras marcadas: heredadas {w(lambda a,b:a==1)}, reales {w(lambda a,b:b==1)}")

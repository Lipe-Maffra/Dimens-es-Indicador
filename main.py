import os, time, traceback
import win32com.client as win32
from win32com.client import constants

caminho_pasta = r'P:\Logística\INDICADORES\2024\ATUALIZAR - INDICADOR\Dimensão'

def listar_planilhas(pasta):
    for nome in os.listdir(pasta):
        if not nome.lower().endswith(('.xlsx', '.xlsm')):
            continue
        if nome.startswith('~$'):
            continue
        yield os.path.join(pasta, nome)

def wait_until_done(excel, wb, poll=0.8, timeout=3600):
    """
    Espera terminar: Power Query, QueryTables, Connections e cálculo.
    Sem sleep fixo; sai quando não tiver nada pendente.
    """
    # 1) Tenta a API nativa p/ consultas assíncronas (Power Query)
    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass  # nem toda versão expõe, seguimos no pooling

    t0 = time.time()
    while True:
        pendencias = []

        # Estado de cálculo
        try:
            if excel.CalculationState != constants.xlDone:
                pendencias.append("calc")
        except Exception:
            pass

        # Conexões (OLEDB/ODBC) com possível flag Refreshing
        try:
            for conn in wb.Connections:
                try:
                    if getattr(conn, "Refreshing", False):
                        pendencias.append(f"conn:{getattr(conn,'Name','?')}")
                except Exception:
                    pass
        except Exception:
            pass

        # QueryTables por planilha
        try:
            for sh in wb.Worksheets:
                try:
                    for qt in sh.QueryTables():
                        try:
                            if getattr(qt, "Refreshing", False):
                                pendencias.append(f"qt:{getattr(qt,'Name','?')}")
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

        # Pivot caches (ocasional)
        try:
            for pc in wb.PivotCaches():
                try:
                    if getattr(pc, "Refreshing", False):
                        pendencias.append("pivotcache")
                except Exception:
                    pass
        except Exception:
            pass

        if not pendencias:
            return  # acabou tudo

        if time.time() - t0 > timeout:
            raise TimeoutError(f"Timeout aguardando refresh: {pendencias}")

        time.sleep(poll)

def atualizar_workbook(excel, arquivo):
    print(f">> Abrindo (oculto): {arquivo}")
    wb = excel.Workbooks.Open(arquivo, ReadOnly=False)

    # Deixa o Excel enxuto e “mudo”
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    try:
        # Força consultas síncronas onde for possível (evita “sair antes da hora”)
        try:
            for conn in wb.Connections:
                try:
                    ole = getattr(conn, "OLEDBConnection", None) or getattr(conn, "ODBCConnection", None)
                    if ole is not None and hasattr(ole, "BackgroundQuery"):
                        ole.BackgroundQuery = False
                except Exception:
                    pass
        except Exception:
            pass

        try:
            for sh in wb.Worksheets:
                try:
                    for qt in sh.QueryTables():
                        try:
                            if hasattr(qt, "BackgroundQuery"):
                                qt.BackgroundQuery = False
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

        # 1) RefreshAll cobre PQ + conexões + pivots
        try:
            wb.RefreshAll()
        except Exception as e:
            print(f"[warn] wb.RefreshAll: {e}")

        # 2) Aguardar terminar (sem sleep fixo)
        wait_until_done(excel, wb)

        # 3) Reforço: Atualizar PivotTables explicitamente
        for sh in wb.Worksheets:
            try:
                for pvt in sh.PivotTables():
                    try:
                        pvt.RefreshTable()
                    except Exception as e:
                        print(f"[warn] Pivot '{getattr(pvt,'Name','?')}': {e}")
            except Exception:
                pass

        # 4) Salvar e fechar
        print("   salvando…")
        wb.Save()
    finally:
        print("<< Fechando")
        try:
            wb.Close(SaveChanges=True)
        except Exception:
            pass
        # Restaura sinalizadores
        try: excel.ScreenUpdating = True
        except Exception: pass
        try: excel.EnableEvents = True
        except Exception: pass

def main():
    if not os.path.exists(caminho_pasta):
        print(f"[erro] Pasta não encontrada: {caminho_pasta}")
        return

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        ok = 0
        for caminho in listar_planilhas(caminho_pasta):
            try:
                atualizar_workbook(excel, caminho)
                ok += 1
            except Exception:
                print("[ERRO NO ARQUIVO]")
                traceback.print_exc()
        print(f"OK: concluído. {ok} arquivo(s) processado(s).")
    finally:
        try: excel.Quit()
        except Exception: pass

if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n[ERRO FATAL]")
        traceback.print_exc()
        input("\nPressione Enter para sair…")

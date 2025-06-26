import pandas as pd
import openpyxl
import os
from datetime import datetime

def filtrar_registros_sem_bairro(arquivo_entrada, coluna_bairro='BAIRRO', arquivo_saida='sem_bairro.xlsx'):
    print(f"🔍 Iniciando análise do arquivo: {arquivo_entrada}")
    print("="*60)
    
    try:
        if not os.path.exists(arquivo_entrada):
            print(f"❌ Erro: Arquivo '{arquivo_entrada}' não encontrado!")
            return False, 0, 0
        
        print("📂 Carregando arquivo Excel...")
        df = pd.read_excel(arquivo_entrada)
        df.columns = df.columns.str.strip()  # ✅ Remove espaços extras dos nomes das colunas
        
        print(f"✅ Arquivo carregado com sucesso!")
        print(f"📊 Total de registros no arquivo: {len(df)}")
        print(f"📋 Colunas encontradas: {list(df.columns)}")
        
        if coluna_bairro not in df.columns:
            print(f"❌ Erro: Coluna '{coluna_bairro}' não encontrada no arquivo!")
            print(f"💡 Colunas disponíveis: {list(df.columns)}")
            return False, len(df), 0
        
        print(f"\n🔎 Analisando coluna '{coluna_bairro}'...")
        total_registros = len(df)

        registros_nao_nulos = df[coluna_bairro].notna().sum()
        registros_nulos = df[coluna_bairro].isna().sum()
        registros_vazios = (df[coluna_bairro] == '').sum()
        registros_espacos = df[coluna_bairro].apply(lambda x: str(x).strip() == '' if pd.notna(x) else False).sum()
        
        print(f"📈 Estatísticas da coluna '{coluna_bairro}':")
        print(f"   • Registros não nulos: {registros_nao_nulos}")
        print(f"   • Registros nulos (NaN/None): {registros_nulos}")
        print(f"   • Registros com string vazia: {registros_vazios}")
        print(f"   • Registros só com espaços: {registros_espacos}")
        
        condicao_sem_bairro = (
            df[coluna_bairro].isna() |
            (df[coluna_bairro] == '') |
            (df[coluna_bairro].astype(str).str.strip() == '')
        )
        
        df_sem_bairro = df[condicao_sem_bairro].copy()
        registros_sem_bairro = len(df_sem_bairro)
        
        print(f"📊 Resultado da filtragem:")
        print(f"   • Total de registros: {total_registros}")
        print(f"   • Registros SEM bairro: {registros_sem_bairro}")
        print(f"   • Registros COM bairro: {total_registros - registros_sem_bairro}")
        
        if registros_sem_bairro == 0:
            print("🎉 Ótima notícia! Não há registros sem bairro no arquivo.")
            return True, total_registros, 0
        
        print(f"\n👁️ Amostra dos registros sem bairro:")
        print("="*50)
        amostra = df_sem_bairro.head(min(5, len(df_sem_bairro)))
        
        for idx, (index, row) in enumerate(amostra.iterrows(), 1):
            print(f"📄 Registro {idx} (linha {index + 2}):")
            colunas_importantes = []
            for col in df.columns:
                if col != coluna_bairro and pd.notna(row[col]) and str(row[col]).strip():
                    colunas_importantes.append(f"{col}: {row[col]}")
            if colunas_importantes:
                print(f"   {' | '.join(colunas_importantes[:3])}")
            valor_bairro = row[coluna_bairro]
            if pd.isna(valor_bairro):
                print(f"   {coluna_bairro}: [NULO]")
            elif valor_bairro == '':
                print(f"   {coluna_bairro}: [VAZIO]")
            else:
                print(f"   {coluna_bairro}: ['{valor_bairro}']")
            print()
        
        if len(df_sem_bairro) > 5:
            print(f"... e mais {len(df_sem_bairro) - 5} registros.")
        
        print(f"\n💾 Salvando arquivo: {arquivo_saida}...")
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            df_sem_bairro.to_excel(writer, sheet_name='Registros Sem Bairro', index=False)
            worksheet = writer.sheets['Registros Sem Bairro']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max(max_length + 2, 10), 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            worksheet.insert_rows(1, 3)
            worksheet['A1'] = f"Registros sem bairro extraídos de: {arquivo_entrada}"
            worksheet['A2'] = f"Data de extração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
            worksheet['A3'] = f"Total de registros sem bairro: {registros_sem_bairro}"
            
            from openpyxl.styles import Font, PatternFill
            for row in range(1, 4):
                worksheet[f'A{row}'].font = Font(bold=True)
                worksheet[f'A{row}'].fill = PatternFill(start_color='E6E6FA', end_color='E6E6FA', fill_type='solid')
        
        print(f"✅ Arquivo salvo com sucesso!")
        print(f"📁 Localização: {os.path.abspath(arquivo_saida)}")
        print(f"📊 Aba criada: 'Registros Sem Bairro'")
        print(f"🔢 Registros salvos: {registros_sem_bairro}")
        
        print(f"\n🎯 RESUMO FINAL:")
        print("="*40)
        print(f"📂 Arquivo origem: {arquivo_entrada}")
        print(f"📂 Arquivo destino: {arquivo_saida}")
        print(f"📊 Total de registros: {total_registros}")
        print(f"❌ Registros sem bairro: {registros_sem_bairro}")
        print(f"✅ Registros com bairro: {total_registros - registros_sem_bairro}")
        print(f"📈 Taxa de completude: {((total_registros - registros_sem_bairro) / total_registros * 100):.1f}%")
        
        return True, total_registros, registros_sem_bairro
        
    except Exception as e:
        print(f"❌ Erro durante o processamento: {str(e)}")
        print(f"💡 Verifique se o arquivo está fechado e tente novamente.")
        return False, 0, 0

# Execução do script
if __name__ == "__main__":
    print("🚀 FILTRADOR DE REGISTROS SEM BAIRRO")
    print("="*50)

    arquivo_entrada = "japeri.xlsx"  # <- Substitua pelo nome correto, se necessário
    arquivo_saida = "sem_bairro.xlsx"
    coluna_bairro = "BAIRRO"

    sucesso, total, sem_bairro = filtrar_registros_sem_bairro(
        arquivo_entrada=arquivo_entrada,
        coluna_bairro=coluna_bairro,
        arquivo_saida=arquivo_saida
    )

    if sucesso:
        if sem_bairro > 0:
            print(f"\n🎉 Processo concluído com sucesso!")
            print(f"📄 Arquivo '{arquivo_saida}' criado com {sem_bairro} registros.")
        else:
            print(f"\n✨ Processo concluído! Nenhum registro sem bairro encontrado.")
    else:
        print(f"\n❌ Processo falhou. Verifique os erros acima.")

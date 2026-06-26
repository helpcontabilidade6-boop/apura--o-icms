import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Apuração Fiscal", layout="wide")

st.title("📊 Apuração Fiscal Mensal")

arquivo_matriz = st.file_uploader(
    "Resumo por CFOP - Matriz",
    type=["xls", "xlsx"],
    key="matriz"
)

arquivo_filial = st.file_uploader(
    "Resumo por CFOP - Filial",
    type=["xls", "xlsx"],
    key="filial"
)

CFOPS_VENDA = [
    5401,
    5403,
    5405,
    6102,
    6108,
    6110,
    6401,
    5102,
    6403
]

CFOPS_DEVOLUCAO = [
    1202,
    1411,
    2411
]


def moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def processar(df, valor_col, imposto_col=8):

    df = df.copy()

    CFOP_COL = 0
    IMPOSTO_COL = imposto_col

    df["cfop_num"] = (
        df[CFOP_COL]
        .astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )

    df["cfop_num"] = pd.to_numeric(
        df["cfop_num"],
        errors="coerce"
    )

    df["valor_num"] = pd.to_numeric(
        df[valor_col],
        errors="coerce"
    ).fillna(0)

    df["imposto_num"] = pd.to_numeric(
        df[IMPOSTO_COL],
        errors="coerce"
    ).fillna(0)

    receita = df.loc[
        df["cfop_num"].isin(CFOPS_VENDA),
        "valor_num"
    ].sum()

    devolucao = df.loc[
        df["cfop_num"].isin(CFOPS_DEVOLUCAO),
        "valor_num"
    ].sum()

    icms = df.loc[
        df["cfop_num"].isin(CFOPS_VENDA),
        "imposto_num"
    ].sum()

    return receita, devolucao, icms


if arquivo_matriz or arquivo_filial:

    try:
        receita_matriz = devol_matriz = icms_matriz = 0
        receita_filial = devol_filial = icms_filial = 0

        if arquivo_matriz:
            matriz = pd.read_excel(
                arquivo_matriz,
                header=None
            )
            receita_matriz, devol_matriz, icms_matriz = processar(
                matriz,
                valor_col=5,
                imposto_col=8
            )        

        if arquivo_filial:
            filial = pd.read_excel(
                arquivo_filial,
                header=None
            )

            receita_filial, devol_filial, icms_filial = processar(
                filial,
                valor_col=6,
                imposto_col=9
            )

        receita_total = receita_matriz + receita_filial
        devol_total = devol_matriz + devol_filial
        icms_total = icms_matriz + icms_filial

        base = (
            receita_total
            - devol_total
            - icms_total
        )

        pis = base * 0.0065
        cofins = base * 0.03

        resultado = pd.DataFrame({
            "Descrição": [
                "Receita Bruta",
                "(-) Devoluções",
                "(-) ICMS",
                "Base PIS/COFINS",
                "PIS 0,65%",
                "COFINS 3%"
            ],
            "Matriz": [
                receita_matriz,
                devol_matriz,
                icms_matriz,
                "",
                "",
                ""
            ],
            "Filial": [
                receita_filial,
                devol_filial,
                icms_filial,
                "",
                "",
                ""
            ],
            "Total": [
                receita_total,
                devol_total,
                icms_total,
                base,
                pis,
                cofins
            ]
        })

        arquivo_excel = BytesIO()

        with pd.ExcelWriter(
            arquivo_excel,
            engine="openpyxl"
        ) as writer:

            resultado.to_excel(
                writer,
                sheet_name="Apuração",
                index=False
            )

        st.download_button(
            "📥 Baixar Apuração.xlsx",
            data=arquivo_excel.getvalue(),
            file_name="Apuracao_PIS_COFINS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("📊 Resultado da Apuração")

        col1, col2, col3 = st.columns(3)

        with col1:

            st.metric(
                "Receita Matriz",
                moeda(receita_matriz)
            )

            st.metric(
                "Devoluções Matriz",
                moeda(devol_matriz)
            )

            st.metric(
                "ICMS Matriz",
                moeda(icms_matriz)
            )

        with col2:

            st.metric(
                "Receita Filial",
                moeda(receita_filial)
            )

            st.metric(
                "Devoluções Filial",
                moeda(devol_filial)
            )

            st.metric(
                "ICMS Filial",
                moeda(icms_filial)
            )

        with col3:

            st.metric(
                "Receita Total",
                moeda(receita_total)
            )

            st.metric(
                "Devoluções Total",
                moeda(devol_total)
            )

            st.metric(
                "ICMS Total",
                moeda(icms_total)
            )

        st.divider()

        st.metric(
            "Base PIS/COFINS",
            moeda(base)
        )

        st.metric(
            "PIS (0,65%)",
            moeda(pis)
        )

        st.metric(
            "COFINS (3%)",
            moeda(cofins)
        )

    except Exception as e:
        st.error(f"Erro: {str(e)}")
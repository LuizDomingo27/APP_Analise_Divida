import streamlit as st
from datetime import datetime


def rodape(empresa="Minha Empresa", ano=None):
    """
    Gera um rodapé para aplicação Streamlit.
    
    Args:
        empresa (str): Nome da empresa ou sistema
        ano (int): Ano do rodapé (opcional, usa o ano atual se não informado)
    """

    ano = datetime.now().year
    texto = f"© {ano} - {empresa}. Todos os direitos reservados."

    # Adiciona linha separadora e texto do rodapé
    st.divider()
    st.markdown(
        f"<div style='text-align: center; font-size: 12px; color: gray;'>{texto}</div>",
        unsafe_allow_html=True
    )
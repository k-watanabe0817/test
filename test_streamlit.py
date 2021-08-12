import streamlit as st
import pandas as pd

def main():
    # タイトル
    st.title('Application title')
    # ヘッダ
    st.header('Header')
    # 純粋なテキスト
    st.text('Some text')
    # サブレベルヘッダ
    st.subheader('Sub header')
    # マークダウンテキスト
    st.markdown('**Markdown is available **')
    # LaTeX テキスト
    st.latex(r'\bar{X} = \frac{1}{N} \sum_{n=1}^{N} x_i')
    # コードスニペット
    st.code('print(\'Hello, World!\')')
    # エラーメッセージ
    st.error('Error message')
    # 警告メッセージ
    st.warning('Warning message')
    # 情報メッセージ
    st.info('Information message')
    # 成功メッセージ
    st.success('Success message')
    # 例外の出力
    st.exception(Exception('Oops!'))
    # 辞書の出力
    d = {
        'foo': 'bar',
        'users': [
            'alice',
            'bob',
            'kazuki',
            'bob2',
        ],
    }
    st.json(d)

if __name__ == '__main__':
    main()
# Otimização Dinâmica de Validades & Dashboard Gerencial

Análise Preventiva Visando Mitigar Perdas de Matérias-Primas.

Solução totalmente desenvolvida internamente utilizando Python, Pandas e Streamlit.                        
Para resolver essa dor, uni meus 6 anos de experiência prática no chão de fábrica da Pesagem com a minha bagagem técnica em Análise de Sistemas, pós-graduação em Back-end e especialização em Data Science.

Como Funcionará:
1° O analista ou assistente recebe a OP para pesagem;
2° Antes de solicitar a MP para pesagem, insere os códigos na aplicação (interface rápida por separação de vírgulas), para validar se alguma MP deve ter preferência;
3° O sistema valida instantaneamente se há lotes MP críticos em estoque para aquela OP.
4° O analista ou assistente solicita a MP específica, evitando deixá-la vencer.

No SAP, exportamos a planilha detalhada de matérias-primas em estoque, pela transação LX02;
Importamos o arquivo xlsx para a aplicação web e recebemos as informações rapidamente, dada a velocidade de processamento de dados do código em Python.
A aplicação analisa os próximos 180 dias (6 meses) e retorna as informações consolidadas que desejamos.

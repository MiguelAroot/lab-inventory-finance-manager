# Controle financeiro de reagentes — Laboratório 217 / DQ-UFMG

O Excel é o documento principal. O Flask apenas facilita o preenchimento, CRUD, upload, reutilização e download.

## Abas
- **Controle**: ID automático, reagente, data, para quem, observação, preço, total pago, saldo, status e forma de compra.
- **Pagamentos**: permite lançar vários pagamentos para o mesmo ID (parcelamento, FAPEMIG etc.).
- **Resumo**: soma compras, pagamentos e dívida.
- **Observações**: anotações adicionais, opcionalmente ligadas a um ID.

## Instalação
No PowerShell, dentro da pasta do projeto:

```powershell
python -m pip install -r requirements.txt
python app.py
```

Se você usa explicitamente o Python 3.14:

```powershell
& C:\Users\(nome_usuario)\AppData\Local\Programs\Python\Python314\python.exe -m pip install -r requirements.txt
& C:\Users\(nome_usuario)\AppData\Local\Programs\Python\Python314\python.exe app.py
```

Abra `http://127.0.0.1:5000`.

## Importante sobre o ID
O usuário não digita o ID. O modelo já possui IDs reservados para as linhas. O site utiliza a primeira linha livre e preserva o ID mesmo se o registro for apagado.

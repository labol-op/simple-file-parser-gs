#!/usr/bin/env python
# coding: utf-8

# In[79]:


import os
import pandas as pd
import gspread
from datetime import datetime


# In[116]:


def formatar_data(data) -> str:
    data_formatada = datetime.strftime(data, format = "%d/%m/%Y")
    return data_formatada

def formatar_hora(hora) -> str:
    hora_formatada = hora.strftime(format = "%H:%M:%S")
    return hora_formatada


# In[15]:


x = pd.read_excel('order_h.XLSX')


# In[17]:


_bopisloja =  x[['Centro', 'Pedido Internet - Bemol On-line', 'Data do Pedido Bemol On-line']]


# In[85]:


bopisloja = _bopisloja.query('Centro > 0')
bopisloja['Data do Pedido Bemol On-line'] = bopisloja['Data do Pedido Bemol On-line'].apply(formatar_data)


# In[126]:


bopishora = x[['Pedido Internet - Bemol On-line','Data do Pedido Bemol On-line',
               'Hora do Pedido Bemol On-line','Caractere 1','Códido de Identificação do site']]
bopishora['Data do Pedido Bemol On-line'] = bopishora['Data do Pedido Bemol On-line'].apply(formatar_data)
bopishora['Hora do Pedido Bemol On-line'] = bopishora['Hora do Pedido Bemol On-line'].apply(formatar_hora)
bopishora.fillna('', inplace = True)


# In[41]:


basepedidos = x[['Pedido Internet - Bemol On-line','Data do Pedido Bemol On-line',
               'Hora do Pedido Bemol On-line','Códido de Identificação do site']]


# In[130]:


basepedidos['Data do Pedido Bemol On-line'] = basepedidos['Data do Pedido Bemol On-line'].apply(formatar_data)
basepedidos['Hora do Pedido Bemol On-line'] = basepedidos['Hora do Pedido Bemol On-line'].apply(formatar_hora)


# In[47]:


fname = os.path.abspath("./test-auto-330418-221ec3d586fa.json")
gc = gspread.service_account(fname)
sh = gc.open("Cópia de Acompanhamento de pedidos SETEMBRO 2021")


# In[49]:


sh_bopisloja = sh.worksheet('BOPIS Recebido Loja')


# In[50]:


#sh_bopisloja.acell('B3').value


# In[137]:


bopisloja_num = bopisloja.shape[0]+4
sh_bopisloja.update(f'A4:C{bopisloja_num}',bopisloja.values.tolist())


# In[96]:


sh_bopishora = sh.worksheet('BOPIS por Hora')


# In[125]:


bopishora_num = bopishora.shape[0]+2
sh_bopishora.update(f'A2:E{bopishora_num}',bopishora.values.tolist())


# In[128]:


sh_basepedidos =  sh.worksheet('Base PEDIDOS h/h')


# In[132]:


basepedidos_num = basepedidos.shape[0]+2
sh_basepedidos.update(f'A2:D{basepedidos_num}',basepedidos.values.tolist())


# In[ ]:





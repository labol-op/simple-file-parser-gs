#!/usr/bin/env python
# coding: utf-8

# In[177]:


import os
import pandas as pd
import gspread
from datetime import datetime, timedelta


# In[187]:


def formatar_data(data) -> str:
    DATA_INICIAL = datetime(1899, 12, 30)
    diferenca_datas = data - DATA_INICIAL
    data_formatada = diferenca_datas.days
    return data_formatada

def formatar_hora(hora) -> str:
    params = {
    'hours': hora.hour,
    'minutes': hora.minute,
    'seconds': hora.second}
    
    hora_formatada = timedelta(**params) / timedelta(days=1)
    return hora_formatada


# In[153]:


def limpar_intervalo(aba,intervalo):
    
    aba.batch_clear([intervalo])
    


# In[154]:


x = pd.read_excel('order_h.XLSX')


# In[155]:


_bopisloja =  x[['Centro', 'Pedido Internet - Bemol On-line', 'Data do Pedido Bemol On-line']]


# In[191]:


bopisloja = _bopisloja.query('Centro > 0')
bopisloja['Data do Pedido Bemol On-line'] = bopisloja['Data do Pedido Bemol On-line'].apply(formatar_data)


# In[192]:


bopishora = x[['Pedido Internet - Bemol On-line','Data do Pedido Bemol On-line',
               'Hora do Pedido Bemol On-line','Caractere 1','Códido de Identificação do site']]
bopishora['Data do Pedido Bemol On-line'] = bopishora['Data do Pedido Bemol On-line'].apply(formatar_data)
bopishora['Hora do Pedido Bemol On-line'] = bopishora['Hora do Pedido Bemol On-line'].apply(formatar_hora)
bopishora.fillna('', inplace = True)


# In[193]:


basepedidos = x[['Pedido Internet - Bemol On-line','Data do Pedido Bemol On-line',
               'Hora do Pedido Bemol On-line','Códido de Identificação do site']]


# In[194]:


basepedidos['Data do Pedido Bemol On-line'] = basepedidos['Data do Pedido Bemol On-line'].apply(formatar_data)
basepedidos['Hora do Pedido Bemol On-line'] = basepedidos['Hora do Pedido Bemol On-line'].apply(formatar_hora)


# In[195]:


fname = os.path.abspath("./test-auto-330418-221ec3d586fa.json")
gc = gspread.service_account(fname)
sh = gc.open("Cópia de Acompanhamento de pedidos SETEMBRO 2021")


# In[196]:


sh_bopisloja = sh.worksheet('BOPIS Recebido Loja')


# In[197]:


#sh_bopisloja.acell('B3').value


# In[201]:


bopisloja_num = bopisloja.shape[0]+4
limpar_intervalo(sh_bopisloja,'A4:C1500')
sh_bopisloja.update(f'A4:C{bopisloja_num}',bopisloja.values.tolist())
sh_bopisloja.format(f'C2:C{bopisloja_num}',
                    {'numberFormat':
                     {'type':
                      "DATE"
                      }
                    })


# In[202]:


sh_bopishora = sh.worksheet('BOPIS por Hora')


# In[203]:


bopishora_num = bopishora.shape[0]+2
limpar_intervalo(sh_bopishora,'A2:E1500')

sh_bopishora.update(f'A2:E{bopishora_num}',bopishora.values.tolist())
sh_bopishora.format(f'C2:C{bopishora_num}',
                    {'numberFormat':
                     {'type':
                      "TIME"
                      }
                    })
sh_bopishora.format(f'B2:B{bopishora_num}',
                    {'numberFormat':
                     {'type':
                      "DATE"
                      }
                    })


# In[204]:


sh_basepedidos =  sh.worksheet('Base PEDIDOS h/h')


# In[206]:


basepedidos_num = basepedidos.shape[0]+2
limpar_intervalo(sh_basepedidos,'A2:D2500')
sh_basepedidos.update(f'A2:D{basepedidos_num}',basepedidos.values.tolist())
sh_basepedidos.format(f'C2:C{basepedidos_num}',
                    {'numberFormat':
                     {'type':
                      "TIME"
                      }
                    })
sh_basepedidos.format(f'B2:B{basepedidos_num}',
                    {'numberFormat':
                     {'type':
                      "DATE"
                      }
                    })


# In[ ]:





# 🔧 Assistência Técnica Teston — Sistema de OS e Comissões

Aplicativo Streamlit para gestão de Ordens de Serviço e cálculo de comissões
dos técnicos da oficina mecânica, baseado nos critérios da planilha Excel.

---

## 📁 Estrutura do Projeto

```
oficina_app/
├── app.py          ← Aplicativo principal Streamlit
├── dados.py        ← Tabelas de valores, técnicos, usuários e persistência
├── calculos.py     ← Motor de cálculo de comissões
├── requirements.txt
└── os_data.json    ← Gerado automaticamente ao registrar OS
```

---

## 🚀 Como Rodar

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## 🔑 Credenciais de Acesso

### Administrador / Supervisor
| Usuário      | Senha     | Perfil    |
|-------------|-----------|-----------|
| `admin`     | admin123  | Admin     |
| `supervisor`| super123  | Supervisor|

### Técnicos (senha padrão: `teston123`)
| Usuário         | Nome              | Nível          |
|----------------|-------------------|----------------|
| `zaqueu`        | Zaqueu            | Técnico Cinco  |
| `marcos.c`      | Marcos C.         | Técnico Cinco  |
| `joao.p`        | João P.           | Técnico Cinco  |
| `tiago`         | Tiago Bardu       | Técnico Quatro |
| `cristiano.s`   | Cristiano dos S.  | Técnico Quatro |
| `cristiano.b`   | Cristiano B.      | Técnico Três   |
| `rodrigo.j`     | Rodrigo J.        | Técnico Três   |
| `loran`         | Loran H.          | Técnico Três   |
| `bruno.h`       | Bruno H.          | Técnico Três   |
| `claudecir`     | Claudecir         | Técnico Três   |
| `rodrigo.m`     | Rodrigo M.        | Técnico Três   |
| `felipe`        | Felipe            | Técnico Dois   |
| `gabriel`       | Gabriel           | Técnico Um     |
| `gustavo`       | Gustavo           | Técnico Um     |
| `vanderlei`     | Vanderlei         | Técnico Um     |
| `jose`          | José              | Técnico Um     |

---

## 📐 Regras de Cálculo (da planilha)

### Tabela de KM Rodado
| Tipo     | R$/km |
|----------|-------|
| Km Um    | R$ 1,70 |
| Km Dois  | R$ 4,10 |
| Km Três  | R$ 5,00 |

### Valor Hora por Nível
| Nível          | R$/hora |
|----------------|---------|
| Técnico Um     | R$ 50   |
| Técnico Dois   | R$ 60   |
| Técnico Três   | R$ 80   |
| Técnico Quatro | R$ 90   |
| Técnico Cinco  | R$ 100  |
| Deslocamento   | R$ 60 (fixo) |
| Munck          | R$ 120 (fixo) |

### Percentual de Comissão por Local
| Local             | % Comissão |
|-------------------|-----------|
| Interno Barracão  | 4%        |
| Campo             | 8%        |
| Deslocamento      | 6%        |
| M.S               | 10%       |

### Fórmula
```
Valor Serviço = Horas × Valor Hora
Valor KM      = KM Rodado × Preço/KM
Valor Munck   = Horas Munck × R$120
Total OS      = Valor Serviço + Valor KM + Valor Munck
Comissão      = Total OS × % do Local
```

---

## 🔄 Fluxo de Trabalho

1. **Técnico faz login** → registra a OS preenchendo todos os campos
2. **Preview automático** da comissão estimada antes de salvar
3. **Supervisor/Admin** visualiza OS pendentes e aprova ou rejeita
4. **Relatório de comissões** disponível por período e por técnico

---

## 🛠️ Personalização

Para ajustar valores (KM, horas, percentuais), edite o arquivo `dados.py`:
- `TABELA_KM` — preços por quilômetro
- `TABELA_HORA` — valor hora por nível técnico
- `PERCENTUAIS_LOCAL` — % de comissão por local

Para adicionar novos técnicos ou usuários, edite os dicionários
`TECNICOS` e `USUARIOS` em `dados.py`.

---

## 📊 Futuros aprimoramentos sugeridos

- [ ] Exportar fechamento mensal em Excel
- [ ] Integração com SharePoint (como os outros apps Teston)
- [ ] Notificações por e-mail ao aprovar/rejeitar
- [ ] Foto/comprovante anexo na OS
- [ ] Histórico de alterações por OS

# Guia de Parâmetros — Planejamento ORT

Documentação dos parâmetros configuráveis via interface do Planejamento ORT.
Fábrica ORT · MP27 / MP28 / MC25 / MC26 · Paraná

---

## 1. Parâmetros — Otimização de Mix

### DIAS_PERÍODO

Horizonte de planejamento em dias. O otimizador aloca demanda e calcula
capacidade para essa janela de tempo.

| Valor | Significado |
|-------|-------------|
| `7`   | Planejar 1 semana |
| `30`  | Planejar 1 mês |

> **Atenção:** deve ser igual ao `DIAS_PERÍODO` do Sequenciamento.
> A interface sincroniza automaticamente ao salvar os parâmetros do Mix.

---

### BASE_DEMANDA_DIAS

Base em que a demanda está cadastrada no Excel. O otimizador escala a
demanda proporcionalmente ao horizonte:

```
demanda_efetiva = demanda_excel × (DIAS_PERÍODO / BASE_DEMANDA_DIAS)
```

| Valor | Significado | Exemplo |
|-------|-------------|---------|
| `365` | Demanda anual | 50.000 t/ano → 958 t em 7 dias |
| `30`  | Demanda mensal | 4.167 t/mês → 972 t em 7 dias |

> **Quando alterar:** sempre que a planilha Excel mudar a convenção de
> cadastro da demanda (anual para mensal ou vice-versa).

---

## 2. Parâmetros — Sequenciamento de Lotes

### LOTE_T

Tamanho fixo de cada lote de produção em toneladas. Todo lote completo
tem exatamente esse volume (exceto o lote parcial final de cada produto).

| Valor | Efeito |
|-------|--------|
| Menor (ex: `300`) | Mais flexibilidade, mais setups, maior perda total de setup |
| Maior (ex: `1000`) | Menos setups, menos flexibilidade, produtos ficam mais concentrados |

**Exemplo:**
Meta de 2.100 t com LOTE_T = 700 → 3 lotes completos.
Meta de 2.100 t com LOTE_T = 300 → 7 lotes completos.

---

### PERDA_SETUP_T_DEFAULT

Toneladas perdidas (fora de especificação) a cada troca de produto.
Usado como valor fixo para MC25/MC26 e como fallback para MP27/MP28
quando a transição não está na matriz de setup.

| Valor | Efeito |
|-------|--------|
| `20`  | Cada troca consome 20 t (padrão) |
| `30`  | Troca mais cara — sequenciador evita ainda mais mudanças |
| `10`  | Troca mais barata — sequenciador aceita mais mudanças |

> **Impacto:** quanto maior o valor, mais o sequenciador tende a agrupar
> lotes do mesmo produto consecutivamente.

---

### FATOR_TOLERANCIA_RITMO

Fração mínima da meta que deve estar cumprida proporcionalmente ao tempo
decorrido. Controla se o sequenciador pode "atrasar" a produção de um
produto.

```
mínimo_acumulado = meta × (dias_decorridos / DIAS_PERÍODO) × FATOR_TOLERANCIA_RITMO
```

| Valor | Efeito |
|-------|--------|
| `0.0` | Desativa o controle de ritmo — sequenciador pode ignorar produtos por ciclos inteiros |
| `0.5` | No meio do horizonte, deve ter pelo menos 50% da meta cumprida (padrão) |
| `1.0` | Ritmo exato — nenhum atraso permitido em nenhum momento |

**Exemplo com meta de 2.100 t e horizonte de 30 dias:**

| Dia | FATOR = 0.5 (mínimo exigido) | FATOR = 1.0 (mínimo exigido) |
|-----|-------------------------------|-------------------------------|
| 7   | 245 t                         | 490 t                         |
| 15  | 525 t                         | 1.050 t                       |
| 30  | 1.050 t                       | 2.100 t                       |

> **Quando aumentar:** se produtos prioritários estão sendo postergados
> demais pelo sequenciador.
> **Quando diminuir:** se o MILP está infeasível por não conseguir cumprir
> o ritmo mínimo em todos os produtos simultaneamente.

---

### FATOR_SPREAD

Teto de quanto pode ser produzido de um produto em um único ciclo, em
relação à sua fatia proporcional do horizonte total.

```
teto_ciclo = meta × (dias_ciclo / DIAS_PERÍODO) × FATOR_SPREAD
```

**Exemplo com meta = 2.100 t, DIAS_PERÍODO = 30, ciclo de 7 dias:**

| FATOR_SPREAD | Teto no ciclo | Interpretação |
|--------------|---------------|---------------|
| `1.0`        | 490 t         | Produz exatamente a fração proporcional |
| `1.5`        | 735 t         | Pode antecipar até 50% a mais |
| `2.0`        | 980 t         | Pode produzir o dobro da fração proporcional |

**Quando usar cada valor:**

| Valor | Situação recomendada |
|-------|----------------------|
| `1.0` | Restrições horárias críticas (Outorga, Evaporação). Distribui carga uniformemente no tempo. |
| `1.2–1.5` | Produtos com prazo concentrado no início do período, ou para liberar máquina mais cedo para outros produtos. |
| `> 1.5` | Cenários de urgência onde um produto precisa ser priorizado num ciclo específico. Aceita risco de violação de restrição. |
| `None` | **Não recomendado para ORT.** Sem teto, o sequenciador concentra toda a produção no primeiro ciclo, quase certamente violando Outorga e Evaporação. |

> **Relação com FATOR_TOLERANCIA_RITMO:** SPREAD define o teto (máximo por
> ciclo), RITMO define o piso (mínimo acumulado). Usar SPREAD = 1.0 e
> RITMO = 0.5 cria uma banda proporcional confortável para o sequenciador.

---

### FRACAO_MIN_LOTE_PARCIAL

Tamanho mínimo do lote parcial (último lote de um produto que fecha o
saldo restante) como fração de LOTE_T. Evita que o sequenciador gere
lotes ínfimos com custo de setup desproporcional.

```
volume_mínimo_parcial = LOTE_T × FRACAO_MIN_LOTE_PARCIAL
```

**Exemplo com LOTE_T = 700:**

| FRACAO | Volume mínimo parcial |
|--------|-----------------------|
| `0.25` | 175 t                 |
| `0.50` | 350 t (padrão)        |
| `0.75` | 525 t                 |

> **Quando diminuir:** se muitos produtos têm saldos pequenos que estão
> sendo descartados pelo sequenciador por não atingirem o piso.
> **Quando aumentar:** se lotes parciais muito pequenos estão gerando
> setups desnecessários no Gantt.

---

### DISPERSAO_UNIFORME_ORT

Toggle (ligado/desligado) que ativa o espaçamento uniforme dos lotes
ao longo dos dias do ciclo. Quando desligado, os lotes são empacotados
sequencialmente desde o início do ciclo.

| Estado | Comportamento |
|--------|---------------|
| `False` | Lotes compactados no início — simples, maximiza utilização contínua |
| `True`  | Lotes espaçados ao longo dos dias — reduz picos de consumo instantâneo |

> Quando `False`, o parâmetro `FATOR_DISPERSAO_ORT` é irrelevante e
> fica desabilitado na interface.

---

### FATOR_DISPERSAO_ORT

Grau de uniformidade do espaçamento quando `DISPERSAO_UNIFORME_ORT`
está ativo. Controla **quando dentro do ciclo** os lotes são posicionados
no calendário.

```
intervalo_entre_lotes = (total_dias_ciclo × FATOR_DISPERSAO) / n_lotes
```

**Exemplo visual com 3 lotes num ciclo de 7 dias:**

```
FATOR = 0.0 (compactado):
Dia 1: [L1][L2][L3]  Dias 2-7: ───────────

FATOR = 0.5 (meio-termo):
Dia 1: [L1][L2]  Dia 4: [L3]  Dias 5-7: ──

FATOR = 1.0 (uniforme):
Dia 1: [L1]  Dia 3: [L2]  Dia 5: [L3]
```

| Valor | Efeito | Usar quando |
|-------|--------|-------------|
| `0.0` | Igual a desativado | Restrições horárias com folga ampla |
| `0.5` | Meio-termo (padrão) | Ponto de equilíbrio geral |
| `1.0` | Totalmente uniforme | Violações de Outorga ou Evaporação persistentes |

---

## 3. Diferença entre FATOR_SPREAD e FATOR_DISPERSAO_ORT

São parâmetros complementares que atuam em camadas diferentes:

| Aspecto | FATOR_SPREAD | FATOR_DISPERSAO_ORT |
|---------|--------------|----------------------|
| **Controla** | Quantidade total produzida no ciclo | Espaçamento entre lotes dentro do ciclo |
| **Atua em** | MILP (decisão de quantos lotes alocar) | Empacotamento (onde no calendário cada lote vai) |
| **Impacto principal** | Restrições agregadas por ciclo | Restrições horárias instantâneas |
| **Analogia** | "Orçamento" de produção | "Como gastar o orçamento ao longo dos dias" |

**Exemplo combinado:**

Meta de 2.100 t, ciclo de 7 dias, LOTE_T = 700:

- SPREAD = 1.0 → MILP aloca 1 lote (700 t) no ciclo
- DISPERSAO = 1.0 → esse lote é posicionado no meio do ciclo (dia 3-4)

- SPREAD = 1.5 → MILP pode alocar até 2 lotes (1.400 t) no ciclo
- DISPERSAO = 0.5 → os 2 lotes ficam no início e meio do ciclo

**Configurações recomendadas por cenário:**

| Cenário | SPREAD | DISPERSAO |
|---------|--------|-----------|
| Conservador (restrições críticas) | `1.0` | `1.0` |
| Equilibrado (padrão) | `1.0` | `0.5` |
| Agressivo (prazo urgente) | `1.5` | `0.5` |
| Máxima utilização (restrições folgadas) | `1.5` | `0.0` |

---

*Gerado automaticamente · Planejamento ORT · Klabin*

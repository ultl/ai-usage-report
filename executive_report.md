# AI Dev Journal — Sprint 0 Executive Report

## 1. Executive Summary
- Across 32 development sessions, AI delivered **115.0 hours of assessed time saved** versus **68.6 hours self-reported**, meaning the pilot captured substantial productivity value that staff did not fully recognize.
- The overall savings rate was **52.6% assessed** versus **49.2% self-reported**. In business terms, users are broadly positive on AI, but they are still **understating the scale of the benefit**.
- The average self-report accuracy on AI savings was uneven: some staff were close to reality, but others recognized only a fraction of AI’s contribution. The most concerning case was **Naduc11 at 30% accuracy**, indicating a major perception gap.
- Quality was strong: the average rating was **4.28/5**, with **15 of 32 sessions rated 5★**, suggesting AI is not only saving time but also producing outputs users value.

![Executive KPI summary](charts_output/03_kpi_summary.png)

## 2. Overall Impact Assessment

| KPI | Assessed (AI) | Self-Reported | Deviation | Business meaning |
|---|---:|---:|---:|---|
| Total Sessions | 32 | 32 | — | Same activity base |
| Staff Count | 6 | 6 | — | Same participant base |
| EST (No AI) | 218.5h | 139.5h | -79.0h | Users estimated the manual workload **79.0h lower** than AI did |
| Actual (With AI) | 103.5h | 70.9h | -32.6h | Users reported less time spent with AI than the objective estimate |
| Hours Saved | 115.0h | 68.6h | -46.4h | Users recognized only part of the value AI delivered |
| Savings % | 52.6% | 49.2% | -3.4 pts | Users slightly understate AI’s productivity lift |

### What the EST gap means
Users estimated **139.5 hours** of manual work, but AI assessed **218.5 hours** — meaning users think their tasks are **36% easier than they objectively are**. This underestimation means users do not fully appreciate the complexity AI is handling for them, especially on higher-complexity work.

### What the Saved gap means
Users reported saving **68.6 hours**, but AI assessed **115.0 hours** of actual savings — users are not recognizing **46.4 hours** of AI-delivered value. In plain terms, the organization is getting more benefit from AI than employees are consciously crediting it for.

**Interpretation:** the pilot shows real productivity gains, but the workforce is systematically conservative in how much value they attribute to AI. That matters because under-recognition can slow adoption, reduce trust in the tool, and weaken the business case if leadership relies only on self-reported numbers.

## 3. Self-Report Accuracy Analysis

![User vs AI comparison](charts_output/09_user_vs_ai_comparison.png)

| Staff | Assessed Saved | Self-Reported Saved | Accuracy | Interpretation |
|---|---:|---:|---:|---|
| Tester | 27.0h | 23.5h | 87% | Perception closely matches reality; Tester understands AI’s contribution well |
| Lnmai2 | 18.0h | 15.0h | 83% | Fairly aligned, but still under-recognizes some AI value |
| Naduc11 | 42.5h | 12.6h | 30% | Major blind spot: perceives only about a third of AI’s actual contribution |
| Nvbinh3 | 11.0h | 9.5h | 86% | Strong alignment; good awareness of AI impact |
| Dngiang | 8.5h | 5.0h | 59% | Moderate under-reporting; sees AI help, but not fully |
| Ltluyen8 | 8.0h | 3.0h | 38% | Low recognition of AI value; likely taking assistance for granted |

**Most accurate reporters:** Tester (87%) and Nvbinh3 (86%).  
**Least accurate reporter:** Naduc11 (30%), followed by Ltluyen8 (38%).

### Behavioral insight
- **Naduc11** thinks tasks are much simpler than they are. They estimated manual work at a level that misses most of the complexity AI is absorbing, so they recognize only **12.6h** of savings versus **42.5h** assessed.
- **Ltluyen8** also underestimates AI’s contribution, recognizing only **3.0h** of savings against **8.0h** assessed.
- **Tester** and **Nvbinh3** are the best calibrated: they appear to understand both the baseline effort and the value AI adds.

### Efficiency anomalies
- **Tester** reported **61.8% efficiency** versus **50.0% assessed** — they are **overestimating their productivity gain**.
- **Nvbinh3** reported **79.2%** versus **45.8% assessed** — a very large inflation of perceived efficiency.
- **Dngiang** and **Ltluyen8** are closer to reality on efficiency, but still understate AI’s savings.

## 4. AI Tool Effectiveness

![EST vs Actual by tool](charts_output/04_est_actual_tool.png)  
![Rating distribution](charts_output/06_rating_distribution.png)

| Tool | Sessions | Assessed Saved | Assessed Savings % | Self-Reported Saved | Avg Rating |
|---|---:|---:|---:|---:|---:|
| Claude | 16 | 59.5h | 53.1% | 33.4h | 4.38/5 |
| ChatGPT | 5 | 18.0h | 55.4% | 6.2h | 4.40/5 |
| Other | 7 | 18.0h | 50.0% | 15.0h | 4.14/5 |
| Claude, Chat GPT | 2 | 8.5h | 60.7% | 5.0h | 4.00/5 |
| Kilo Code | 1 | 5.0h | 50.0% | 6.0h | 4.00/5 |
| GitHub Copilot | 1 | 6.0h | 42.9% | 3.0h | 4.00/5 |

**What this means:**
- **Claude** delivered the largest absolute value: **59.5 hours saved across 16 sessions**, making it the most important tool in the pilot.
- **ChatGPT** had the highest assessed savings rate at **55.4%**, slightly ahead of Claude.
- Satisfaction is consistently high across tools, ranging from **4.0 to 4.4/5**, which supports continued adoption.
- The gap between assessed and self-reported savings is largest on **Claude** and **ChatGPT**, suggesting users may be benefiting more than they realize.

## 5. Category Analysis

![EST vs Actual by category](charts_output/05_est_actual_category.png)

| Category | Sessions | Assessed Saved | Assessed % | Self-Reported Saved | Gap |
|---|---:|---:|---:|---:|---:|
| Documentation | 12 | 37.0h | 52.9% | 29.0h | -8.0h |
| Backend | 3 | 10.5h | 43.8% | 9.0h | -1.5h |
| Refactor | 2 | 7.0h | 50.0% | 8.0h | +1.0h |
| Database Design | 3 | 11.5h | 56.1% | 4.1h | -7.4h |
| Business Logic | 3 | 12.5h | 59.5% | 3.6h | -8.9h |
| Integrated Security | 2 | 12.0h | 54.5% | 2.7h | -9.3h |
| Integration Design | 2 | 6.5h | 59.1% | 2.2h | -4.3h |

**Key findings:**
- **Documentation** is the largest use case by volume: **12 sessions** and **37.0h saved** assessed. This is where AI is already delivering repeatable value.
- The biggest perception gaps are in **Integrated Security (-9.3h)**, **Business Logic (-8.9h)**, and **Documentation (-8.0h)**.
- In **Database Design**, users reported saving only **4.1h** while AI assessed **11.5h** — users drastically underestimate AI’s contribution to complex data modeling tasks, possibly because they attribute the quality of AI output to their own domain knowledge rather than the tool.
- **Refactor** is the only category where self-report slightly exceeds assessed savings, suggesting users may over-credit AI for straightforward cleanup work.

## 6. SDLC Stage Distribution

![SDLC stage breakdown](charts_output/01_sdlc_tasks_by_stage.png)

| Stage | Task Count | Assessed % | Self-Reported % | Interpretation |
|---|---:|---:|---:|---|
| Planning / Requirements | 1 | 50.0% | 90.0% | Users greatly overestimate AI’s productivity in a single planning task |
| Design / Architecture | 9 | 54.1% | 42.4% | Users undervalue AI’s help in design-heavy work |
| Development / Implementation | 3 | 46.7% | 54.5% | Users slightly overestimate AI’s impact here |
| Testing / QA | 1 | 60.0% | 40.0% | Users under-recognize AI’s value in QA |
| Debugging / Bug Fix | 2 | 50.0% | 60.0% | Users think AI helps more than assessed |
| Refactoring / Code Quality | 2 | 50.0% | 61.5% | Users overstate AI’s contribution |
| Deployment / Release | 1 | 56.2% | 44.0% | Users understate AI’s value |
| Operations / Maintenance | 2 | 50.0% | 41.7% | Users understate AI’s value |
| Documentation | 9 | 52.7% | 47.3% | Fairly close, but still conservative |
| Research / Learning | 2 | 59.4% | 38.5% | Users significantly under-recognize AI’s help |

**Interpretation:** the largest mismatch is in **Planning / Requirements** and **Research / Learning**, where users either overestimate or underestimate AI’s role depending on the task type. This suggests AI is not yet being used consistently as a thinking partner across the lifecycle.

## 7. Prompt Engineering Quality

![Top prompt errors](charts_output/07_top_errors.png)  
![Error heatmap](charts_output/08_error_heatmap.png)

The most common prompt issues were:
- **Clear and Format**: 28 occurrences
- **Missing Context**: 11
- **Tool or Environment Missing**: 6
- **Ambiguous Scope**: 4

**What this means:** most productivity loss is not from the AI model itself, but from weak prompting discipline. Staff are often asking for outputs without enough structure, context, or constraints, which increases rework and reduces the perceived value of AI.

**Recommendation from the data:** standardize prompt templates with required sections for goal, context, input data, output format, and acceptance criteria.

## 8. Recommendations
1. **Roll out prompt templates immediately.** The top issue, **Clear and Format (28 cases)**, shows that the biggest gain is not a new model but better instructions.
2. **Target training on high-gap categories.** Focus on **Integrated Security (12.0h assessed saved vs 2.7h self-reported)** and **Business Logic (12.5h vs 3.6h)**, where users are missing most of AI’s value.
3. **Use Claude as the default pilot tool for broad adoption.** It delivered **59.5h saved across 16 sessions** and the highest satisfaction at **4.38/5**.
4. **Coach low-accuracy users individually.** **Naduc11 (30% accuracy)** and **Ltluyen8 (38%)** need support to better understand both task complexity and AI contribution.
5. **Track objective savings in future pilots.** Self-report alone understates value by **46.4h**, so leadership should use assessed metrics as the primary ROI measure.

## 9. Methodology Note
“Assessed (AI)” is an independent blind estimate of manual effort, actual effort with AI, and resulting savings. It is generated without seeing the user’s own numbers, so it serves as the objective baseline. “Self-Reported” reflects the staff member’s own estimate. The gap between the two reveals whether users underestimate task complexity, under-recognize AI’s contribution, or overstate their own productivity gains.
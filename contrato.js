// Fonte única do contrato da MG Locações. Usado pelo app (gera o contrato no celular)
// e pelo script gera_contrato_docx.py (gera o modelo Word). Não edite um sem o outro.
// Marcações: #T título | #S subtítulo | #H cláusula | #Q quadro-resumo | #A assinaturas
// #P quebra de página | [op]...[/op] só com operador | [sop]...[/sop] só sem operador
// {{campo}} = dado do quadro-resumo. **texto** = negrito.

/*CAMPOS*/
var CAMPOS = [
  {"id":"numero","rot":"Contrato nº","grupo":"Contrato","ex":"MG-2026-001"},
  {"id":"cidade_data","rot":"Local e data","grupo":"Contrato","ex":"Campina Grande/PB, 22 de novembro de 2026"},
  {"id":"loc_rep","rot":"Representante da LOCADORA","grupo":"Locadora","ex":""},
  {"id":"loc_rep_cpf","rot":"CPF do representante da LOCADORA","grupo":"Locadora","ex":""},
  {"id":"cli_razao","rot":"LOCATÁRIA (razão social)","grupo":"Locatária","ex":""},
  {"id":"cli_cnpj","rot":"CNPJ/CPF da LOCATÁRIA","grupo":"Locatária","ex":""},
  {"id":"cli_end","rot":"Endereço da LOCATÁRIA","grupo":"Locatária","ex":""},
  {"id":"cli_rep","rot":"Representante da LOCATÁRIA","grupo":"Locatária","ex":""},
  {"id":"cli_rep_cpf","rot":"CPF do representante da LOCATÁRIA","grupo":"Locatária","ex":""},
  {"id":"cli_email","rot":"E-mail para comunicações e cobrança","grupo":"Locatária","ex":""},
  {"id":"cli_whats","rot":"WhatsApp para comunicações e cobrança","grupo":"Locatária","ex":""},
  {"id":"prepostos","rot":"Prepostos autorizados (nome, função, telefone)","grupo":"Locatária","ex":""},
  {"id":"fia_nome","rot":"Fiador e devedor solidário","grupo":"Garantias","ex":""},
  {"id":"fia_cpf","rot":"CPF do fiador","grupo":"Garantias","ex":""},
  {"id":"fia_end","rot":"Endereço do fiador","grupo":"Garantias","ex":""},
  {"id":"fia_conjuge","rot":"Estado civil do fiador e cônjuge anuente (nome e CPF)","grupo":"Garantias","ex":""},
  {"id":"caucao","rot":"Caução em dinheiro","grupo":"Garantias","ex":"R$ 35.000,00 (trinta e cinco mil reais)"},
  {"id":"equip","rot":"Equipamento (tipo, marca, modelo)","grupo":"Equipamento","ex":""},
  {"id":"equip_serie","rot":"Nº de série / chassi e ano","grupo":"Equipamento","ex":""},
  {"id":"acessorios","rot":"Acessórios entregues","grupo":"Equipamento","ex":""},
  {"id":"valor_reposicao","rot":"Valor de reposição do equipamento","grupo":"Equipamento","ex":""},
  {"id":"local_obra","rot":"Local de uso (obra/endereço)","grupo":"Uso","ex":""},
  {"id":"atividade","rot":"Atividade declarada","grupo":"Uso","ex":"Escavação e desmonte de rocha com rompedor"},
  {"id":"prazo_meses","rot":"Prazo mínimo (meses)","grupo":"Prazo e preço","ex":"3 (três)"},
  {"id":"data_inicio","rot":"Início previsto","grupo":"Prazo e preço","ex":""},
  {"id":"data_fim","rot":"Término do prazo mínimo","grupo":"Prazo e preço","ex":""},
  {"id":"valor_mensal","rot":"Aluguel mensal mínimo","grupo":"Prazo e preço","ex":"R$ 35.000,00 (trinta e cinco mil reais)"},
  {"id":"franquia_horas","rot":"Franquia de horas por mês","grupo":"Prazo e preço","ex":"200 (duzentas) horas de horímetro"},
  {"id":"valor_hora_extra","rot":"Valor da hora excedente","grupo":"Prazo e preço","ex":"R$ 175,00 (cento e setenta e cinco reais)"},
  {"id":"adicional_fds","rot":"Acréscimo em sábados, domingos e feriados","grupo":"Prazo e preço","ex":"50% (cinquenta por cento) sobre a hora excedente"},
  {"id":"diaria_operador","rot":"Diária do operador por dia trabalhado","grupo":"Prazo e preço","ex":"R$ 150,00 (cento e cinquenta reais)"},
  {"id":"jornada","rot":"Jornada do operador","grupo":"Prazo e preço","ex":"segunda a sexta-feira, 7h às 17h, com 1h de intervalo"},
  {"id":"mobilizacao","rot":"Mobilização e desmobilização (transporte)","grupo":"Prazo e preço","ex":"por conta da LOCATÁRIA, R$ ____ por trecho"},
  {"id":"dia_vencimento","rot":"Vencimento","grupo":"Pagamento","ex":"dia 10 (dez) do mês seguinte ao período medido"},
  {"id":"dados_bancarios","rot":"Conta para pagamento","grupo":"Pagamento","ex":""}
];
/*FIM*/

var CONTRATO = `#T CONTRATO DE LOCAÇÃO DE EQUIPAMENTO PESADO
[op]
#S COM OPERADOR
[/op]
[sop]
#S SEM OPERADOR
[/sop]
#S Contrato nº {{numero}}

#Q

#H CLÁUSULA 1ª. DAS PARTES E DO QUADRO-RESUMO
**LOCADORA:** MG LOCAÇÕES LTDA, pessoa jurídica de direito privado, inscrita no CNPJ sob o nº 20.441.510/0001-01, com sede na Rua Raimundo da Silva Ribeiro, nº 445, Tambor, Campina Grande/PB, CEP 58414-650, neste ato representada por {{loc_rep}}, CPF nº {{loc_rep_cpf}}.
**LOCATÁRIA:** {{cli_razao}}, inscrita no CNPJ/CPF sob o nº {{cli_cnpj}}, com sede em {{cli_end}}, neste ato representada por {{cli_rep}}, CPF nº {{cli_rep_cpf}}.
**FIADOR E DEVEDOR SOLIDÁRIO:** {{fia_nome}}, CPF nº {{fia_cpf}}, residente em {{fia_end}}, estado civil e cônjuge: {{fia_conjuge}}.
1.1. O Quadro-Resumo acima integra este contrato. Havendo divergência, prevalecem os valores, datas e dados do Quadro-Resumo e, em tudo o mais, as cláusulas gerais.
1.2. As partes declaram que este é um contrato empresarial, paritário e simétrico, negociado livremente, e que cada cláusula foi lida e compreendida, aplicando-se o art. 421-A do Código Civil.
1.3. A LOCATÁRIA declara que usará o EQUIPAMENTO como insumo de sua atividade empresarial ou profissional, e não como destinatária final.

#H CLÁUSULA 2ª. DO OBJETO
2.1. A LOCADORA dá em locação à LOCATÁRIA o equipamento descrito no Quadro-Resumo, com os acessórios ali listados (o "EQUIPAMENTO"), para uso exclusivo no local de uso e na atividade declarados no Quadro-Resumo.
[op]
2.2. A locação inclui operador habilitado, vinculado à LOCADORA, nos termos da Cláusula 12ª.
[/op]
[sop]
2.2. A locação é feita sem operador. A operação, a guarda e a direção dos serviços cabem inteiramente à LOCATÁRIA, nos termos da Cláusula 12ª.
[/sop]
2.3. É proibido à LOCATÁRIA, sem autorização escrita da LOCADORA: retirar o EQUIPAMENTO do local de uso; sublocá-lo, cedê-lo, emprestá-lo ou entregá-lo a terceiros, inclusive a empresas do mesmo grupo ou consórcio; usá-lo em atividade diversa da declarada; e oferecê-lo em garantia de qualquer obrigação.
2.4. O EQUIPAMENTO é e continua sendo da LOCADORA ou de sua legítima posse. Ele não responde por dívidas da LOCATÁRIA, que deverá informar essa condição a qualquer oficial de justiça, autoridade ou credor que pretenda apreendê-lo, comunicando a LOCADORA no mesmo dia.
2.5. A LOCADORA pode instalar e manter no EQUIPAMENTO rastreador, telemetria e câmeras. A LOCATÁRIA não pode desligar, remover ou obstruir esses dispositivos, e os registros por eles gerados servem de prova de localização, horas trabalhadas e condições de uso.

#H CLÁUSULA 3ª. DO PRAZO
3.1. A locação vigora pelo prazo mínimo indicado no Quadro-Resumo, contado da data do Termo de Entrega (Anexo I), que registra a chegada do EQUIPAMENTO ao local de uso.
3.2. Vencido o prazo mínimo sem devolução, a locação continua por prazo indeterminado, mês a mês, nas mesmas condições, podendo qualquer das partes encerrá-la com aviso escrito de 15 (quinze) dias.
3.3. Nos meses de início e de término, o aluguel é cobrado proporcionalmente aos dias de locação, respeitado o faturamento mínimo de 15 (quinze) dias.
3.4. O aluguel corre até a efetiva devolução do EQUIPAMENTO à LOCADORA, com assinatura do Termo de Devolução (Anexo II), ou até a data em que a LOCADORA o retirar nos termos da Cláusula 8ª.

#H CLÁUSULA 4ª. DO PREÇO, DA FRANQUIA E DAS HORAS EXCEDENTES
4.1. A LOCATÁRIA pagará o aluguel mensal mínimo indicado no Quadro-Resumo, que remunera a disponibilidade do EQUIPAMENTO e dá direito à franquia mensal de horas ali prevista.
4.2. O aluguel mensal mínimo é devido por inteiro ainda que a LOCATÁRIA utilize menos horas que a franquia ou paralise os serviços por motivo que não seja defeito do EQUIPAMENTO imputável à LOCADORA, tais como chuva, falta de frente de serviço, falta de combustível, embargo da obra, falta de pagamento por parte do contratante da LOCATÁRIA ou de órgão público, ou decisão interna de reduzir o ritmo da obra.
4.3. As horas que ultrapassarem a franquia, apuradas pelo horímetro, serão cobradas pelo valor da hora excedente previsto no Quadro-Resumo.
4.4. O trabalho em sábados, domingos e feriados depende de solicitação prévia da LOCATÁRIA, por escrito ou por mensagem de seus prepostos, e será cobrado como hora excedente, com o acréscimo previsto no Quadro-Resumo.
[op]
4.5. Correm por conta da LOCATÁRIA a alimentação, o alojamento e o deslocamento local do operador, além da diária por dia trabalhado prevista no Quadro-Resumo.
[/op]
4.6. Os valores serão reajustados a cada 12 (doze) meses pela variação acumulada do IPCA/IBGE, ou por índice oficial que o substitua.

#H CLÁUSULA 5ª. DA MEDIÇÃO E DO FATURAMENTO
5.1. A medição é mensal e se fecha no último dia de cada mês, ou na data de devolução, com base no horímetro, nos apontamentos diários e, quando houver, nos registros de telemetria.
5.2. A LOCADORA enviará a medição e a fatura ao e-mail ou ao WhatsApp indicados no Quadro-Resumo. A LOCATÁRIA terá 3 (três) dias úteis para aprová-la ou impugná-la por escrito, indicando de forma específica o item e o motivo da discordância. Sem impugnação nesse prazo, a medição fica aprovada para todos os efeitos, inclusive para cobrança judicial.
5.3. A impugnação de um item não suspende o pagamento da parte não impugnada, e nunca alcança o aluguel mensal mínimo.
5.4. O pagamento não depende da emissão de boletim de medição pelo sistema interno da LOCATÁRIA, de aprovação de diretoria, de assinatura de contratos com terceiros, de ingresso de investidores, de repasse de órgão público ou de liberação de financiamento. Esses fatos são riscos do negócio da LOCATÁRIA e não podem ser opostos à LOCADORA.
5.5. Nenhum desconto, retenção ou compensação poderá ser feito pela LOCATÁRIA sem acordo escrito prévio e comprovante da despesa.

#H CLÁUSULA 6ª. DO PAGAMENTO
6.1. O vencimento e a conta para pagamento são os do Quadro-Resumo. O pagamento só se considera feito com o crédito na conta da LOCADORA.
6.2. Pagamentos feitos a pessoas não autorizadas por escrito pela LOCADORA, inclusive ao operador, não quitam a dívida.
6.3. A LOCADORA pode emitir boleto, nota de débito ou fatura de locação, e ceder seus créditos a terceiros, independentemente de anuência da LOCATÁRIA.

#H CLÁUSULA 7ª. DAS GARANTIAS
7.1. **Caução.** A LOCATÁRIA depositará, antes da mobilização, a caução em dinheiro prevista no Quadro-Resumo. A LOCADORA poderá usá-la para cobrir qualquer débito vencido, e a LOCATÁRIA deverá recompô-la em 5 (cinco) dias. A caução, sem juros, será devolvida em até 10 (dez) dias após a devolução do EQUIPAMENTO e a quitação de todas as obrigações, descontados os débitos existentes.
7.2. **Fiança e solidariedade.** O FIADOR assina este contrato como devedor solidário e principal pagador de todas as obrigações da LOCATÁRIA, inclusive aluguéis, horas excedentes, encargos de mora, multas, indenizações por danos ou perda do EQUIPAMENTO e despesas de cobrança, até a efetiva devolução do EQUIPAMENTO e a quitação integral, e renuncia ao benefício de ordem (art. 828, I, do Código Civil).
7.2.1. A fiança se prorroga automaticamente com a prorrogação da locação prevista na Cláusula 3.2. O FIADOR só se exonera por notificação à LOCADORA, na forma do art. 835 do Código Civil, e continua responsável por todas as obrigações até 60 (sessenta) dias depois de recebida a notificação. Nesse caso, a LOCATÁRIA apresentará novo fiador ou caução equivalente em 15 (quinze) dias, sob pena de resolução do contrato.
7.2.2. O reajuste, a cobrança de horas excedentes e a prorrogação previstos neste contrato não são aditamento. Qualquer outra alteração das obrigações deverá ser assinada também pelo FIADOR.
7.2.3. Sendo o FIADOR casado ou em união estável, salvo no regime da separação absoluta de bens, o cônjuge ou companheiro assina este contrato para autorizar a fiança (art. 1.647, III, do Código Civil).
7.3. **Consórcio e grupo.** Se a LOCATÁRIA for consórcio, todas as empresas consorciadas assinam este contrato e respondem solidariamente por todas as obrigações dele decorrentes, nos termos do art. 265 do Código Civil e do art. 278, § 1º, da Lei nº 6.404/1976. A empresa que emitir medição, receber fatura ou se beneficiar do EQUIPAMENTO responde solidariamente com a LOCATÁRIA.
7.4. **Título executivo.** As partes reconhecem este contrato como título executivo extrajudicial (art. 784, III e § 4º, do Código de Processo Civil), pelo qual são líquidos, certos e exigíveis o aluguel mensal mínimo, a mobilização e desmobilização, as multas e os encargos aqui previstos. As horas excedentes se provam pelas medições aprovadas, expressa ou tacitamente, na forma da Cláusula 5ª.

#H CLÁUSULA 8ª. DA MORA, DA SUSPENSÃO E DA RETOMADA DO EQUIPAMENTO
8.1. O atraso no pagamento sujeita a LOCATÁRIA, independentemente de aviso, a: multa moratória de 10% (dez por cento) sobre o valor em atraso; juros de mora de 1% (um por cento) ao mês, calculados dia a dia; correção monetária pelo IPCA/IBGE (art. 389, parágrafo único, do Código Civil); e, havendo cobrança por advogado, honorários de 10% (dez por cento) na fase extrajudicial e de 20% (vinte por cento) na judicial (arts. 389 e 395 do Código Civil), sem prejuízo dos honorários de sucumbência. Se a relação vier a ser reconhecida como de consumo, a multa moratória fica reduzida a 2% (dois por cento).
8.2. Com atraso superior a 5 (cinco) dias, a LOCADORA poderá, após aviso por e-mail ou WhatsApp com 48 (quarenta e oito) horas de antecedência, suspender a operação, recolher o operador, se houver, e manter o EQUIPAMENTO parado e travado no canteiro até a regularização (art. 476 do Código Civil). O travamento remoto só será acionado com o EQUIPAMENTO parado e sem risco a pessoas ou à obra. O aluguel continua correndo durante a suspensão, pois o EQUIPAMENTO segue alocado à LOCATÁRIA e indisponível para outros clientes.
8.3. Com atraso superior a 15 (quinze) dias, ou diante de violação das Cláusulas 2ª, 10ª ou 11ª, o contrato poderá ser resolvido de pleno direito pela LOCADORA, mediante simples comunicação escrita, nos termos do art. 474 do Código Civil. Nesse caso:
a) a LOCATÁRIA deve restituir o EQUIPAMENTO em 24 (vinte e quatro) horas, franqueando à LOCADORA, a seus prepostos e transportadores o acesso ao canteiro, com apoio para carregamento;
b) a LOCATÁRIA autoriza desde já o ingresso da LOCADORA no canteiro para retirar o EQUIPAMENTO de forma pacífica, correndo os custos de desmobilização por sua conta. Havendo qualquer oposição no local, a LOCADORA não forçará a retirada e pedirá a reintegração de posse;
c) a recusa ou o embaraço à retirada configura esbulho e autoriza a reintegração de posse com pedido liminar, e a LOCATÁRIA pagará, enquanto não restituir o EQUIPAMENTO, aluguel-pena correspondente ao dobro do aluguel diário proporcional (art. 575 do Código Civil);
d) incide a multa compensatória da Cláusula 15ª.
8.4. A LOCATÁRIA autoriza, após comunicação prévia, o protesto deste contrato, das faturas e das medições (art. 1º da Lei nº 9.492/1997) e a inscrição do débito em cadastros de proteção ao crédito.
8.5. O inadimplemento perante a LOCADORA em qualquer outro contrato também autoriza as medidas desta cláusula.

#H CLÁUSULA 9ª. DA ENTREGA, DA VISTORIA E DA DEVOLUÇÃO
9.1. Na entrega será lavrado o Termo de Entrega e Vistoria (Anexo I), com fotos e vídeo datados, leitura do horímetro, nível de combustível e estado de dentes, unhas, ponteira do rompedor, mangueiras, vidros, fechaduras, tampa do tanque, baterias, esteiras ou pneus e cabine. O preposto que receber o EQUIPAMENTO no canteiro é considerado autorizado pela LOCATÁRIA.
9.2. Se a LOCATÁRIA não assinar o Termo de Entrega em 24 (vinte e quatro) horas, prevalecem as fotos, vídeos e registros enviados pela LOCADORA por WhatsApp ou e-mail.
9.3. O EQUIPAMENTO será devolvido no local de uso, limpo, abastecido no mesmo nível da entrega e no mesmo estado de conservação, ressalvado apenas o desgaste natural compatível com as horas trabalhadas. A vistoria de devolução será feita em até 48 (quarenta e oito) horas, com o Termo de Devolução (Anexo II).
9.4. Os danos constatados serão orçados por concessionária ou oficina especializada. A LOCATÁRIA terá 3 (três) dias úteis para contestar o orçamento de forma fundamentada; sem contestação, o valor fica aprovado e será cobrado na fatura seguinte ou executado como dívida líquida.

#H CLÁUSULA 10ª. DO USO, DA CONSERVAÇÃO E DA INTEGRIDADE DO EQUIPAMENTO
10.1. A LOCATÁRIA usará o EQUIPAMENTO como se fosse seu (art. 569, I, do Código Civil) e se obriga a:
a) usá-lo apenas na atividade declarada, dentro dos limites de capacidade e das instruções do fabricante, sendo vedado usar o rompedor como alavanca, em vazio ou por tempo contínuo acima do recomendado;
b) manter vigilância no canteiro 24 (vinte e quatro) horas por dia e guardar o EQUIPAMENTO, fora do horário de trabalho, em local seguro, plano e iluminado;
c) abastecê-lo somente com diesel S10 de procedência comprovada e filtrado, respondendo pelos danos causados por combustível contaminado ou inadequado;
d) fazer a inspeção diária, o engraxamento e a verificação de níveis antes de cada turno, registrando no apontamento diário;
e) comunicar à LOCADORA, em até 2 (duas) horas, qualquer falha, ruído, vazamento, alarme ou acidente, e interromper a operação imediatamente, sendo vedado prosseguir com o EQUIPAMENTO em falha;
f) não permitir reparo, troca de peças ou intervenção de terceiros sem autorização escrita da LOCADORA;
g) permitir à LOCADORA o acesso ao EQUIPAMENTO, em horário comercial, para inspeção e manutenção.
10.2. A manutenção preventiva programada pelo horímetro é feita pela LOCADORA, que poderá parar o EQUIPAMENTO pelo tempo necessário, até 8 (oito) horas por mês, sem abatimento do aluguel.
10.3. Correm por conta da LOCATÁRIA as peças de desgaste acelerado consumidas na obra: dentes, unhas e adaptadores da caçamba, ponteira, buchas e pinos do rompedor, lâminas de corte e filtros trocados por contaminação. A LOCADORA pode fornecê-las e cobrar o custo comprovado na fatura seguinte.
10.4. A paralisação por defeito do EQUIPAMENTO que não decorra de mau uso deve ser comunicada por escrito. A LOCADORA terá 48 (quarenta e oito) horas úteis para repará-lo ou substituí-lo por equipamento equivalente. Somente depois desse prazo haverá abatimento proporcional dos dias parados.
10.5. Não geram abatimento as paralisações causadas pela LOCATÁRIA, como falta de combustível, de frente de serviço ou de vigilância, ordem de serviço inadequada, falta de pagamento ou suspensão prevista na Cláusula 8ª.
10.6. A responsabilidade da LOCADORA por paralisação, defeito ou falta de operador limita-se ao abatimento proporcional previsto nas Cláusulas 10.4 e 12.5. A LOCADORA não responde por lucros cessantes, atraso de obra, multas ou penalidades que a LOCATÁRIA sofra perante seus contratantes, nem por danos indiretos, salvo dolo ou culpa grave.

#H CLÁUSULA 11ª. DOS DANOS, FURTO, ROUBO E SINISTRO
11.1. Da entrega até a devolução, a LOCATÁRIA responde pelos danos ao EQUIPAMENTO e pela sua perda, inclusive por mau uso, imperícia, negligência, acidente, tombamento, queda, incêndio, alagamento, vandalismo, furto, roubo ou apropriação indébita, e assume expressamente os riscos de caso fortuito e força maior ocorridos no canteiro ou durante o uso (art. 393 do Código Civil). Fica ressalvado o dano causado exclusivamente por defeito do EQUIPAMENTO, cuja prova cabe à LOCATÁRIA.
11.2. Em caso de perda total, furto ou roubo, a LOCATÁRIA pagará à LOCADORA, em 10 (dez) dias, o valor de reposição previsto no Quadro-Resumo, corrigido pelo IPCA desde a assinatura, ou o valor de mercado de equipamento equivalente, o que for maior, além do aluguel mensal mínimo até a data do pagamento da indenização, limitado a 6 (seis) meses.
11.3. Ocorrido furto, roubo, acidente ou dano relevante, a LOCATÁRIA registrará boletim de ocorrência em até 24 (vinte e quatro) horas e entregará cópia à LOCADORA.
11.4. A existência de seguro, contratado por qualquer das partes, não limita a responsabilidade da LOCATÁRIA, que pagará a franquia, os valores não cobertos e os aluguéis do período de reparo.
11.5. As multas, autuações e penalidades de trânsito, ambientais ou administrativas relativas ao período da locação são de responsabilidade da LOCATÁRIA.

[op]
#H CLÁUSULA 12ª. DO OPERADOR
12.1. O operador é contratado e remunerado pela LOCADORA, que responde por seu salário, encargos trabalhistas e previdenciários e pelos equipamentos de proteção individual básicos. Não há vínculo de emprego entre o operador e a LOCATÁRIA.
12.2. Cabe à LOCATÁRIA a direção técnica dos serviços, a sinalização, o isolamento da área, a segurança do canteiro e a integração do operador a seu programa de segurança (NR-1, NR-12 e NR-18), além de fornecer água potável, instalações sanitárias, alimentação e alojamento dignos.
12.3. A LOCATÁRIA não pode exigir do operador trabalho em condição insegura, fora das especificações do EQUIPAMENTO, fora da jornada prevista no Quadro-Resumo sem ajuste prévio, nem serviço de outra natureza. O operador pode recusar ordem que exponha a risco a si, a terceiros ou ao EQUIPAMENTO.
12.4. A LOCATÁRIA responde pelos danos ao EQUIPAMENTO, ao operador e a terceiros decorrentes de suas ordens, da condição do terreno, da falta de sinalização ou de informação sobre interferências (redes, cabos, tubulações, taludes instáveis).
12.5. Faltando o operador por motivo imputável à LOCADORA, ela o substituirá em até 2 (dois) dias úteis; após esse prazo, os dias sem operação serão abatidos proporcionalmente.
12.6. Pagamentos, gratificações ou adiantamentos feitos diretamente ao operador não quitam nenhum valor devido à LOCADORA.
12.7. Para fins fiscais, a parcela correspondente à mão de obra do operador poderá ser faturada em documento próprio, conforme orientação contábil da LOCADORA, sem alteração do valor total.
12.8. Durante a locação e até 6 (seis) meses após o seu término, a LOCATÁRIA não contratará, direta ou indiretamente, o operador da LOCADORA, sob pena de multa equivalente a 1 (um) aluguel mensal mínimo.
[/op]
[sop]
#H CLÁUSULA 12ª. DA OPERAÇÃO PELA LOCATÁRIA
12.1. A LOCATÁRIA operará o EQUIPAMENTO exclusivamente por operador habilitado, com treinamento específico para o modelo (NR-12), vinculado a ela, cujo nome e documento serão informados à LOCADORA antes do início.
12.2. A LOCATÁRIA responde exclusivamente, nas esferas civil, trabalhista, previdenciária, acidentária, ambiental e criminal, pela operação, pelo operador e pelos danos causados a terceiros, ao meio ambiente e ao próprio EQUIPAMENTO.
12.3. Se a LOCADORA for demandada por fato ligado à operação ou ao operador, a LOCATÁRIA deverá assumir a defesa, requerer a exclusão da LOCADORA do processo e reembolsá-la integralmente de condenações, acordos, custas e honorários, em até 5 (cinco) dias do pedido.
[/sop]

#H CLÁUSULA 13ª. DA MOBILIZAÇÃO E DA DESMOBILIZAÇÃO
13.1. O transporte de ida e volta do EQUIPAMENTO (prancha ou carreta), escolta, licenças, carga e descarga correm por conta da LOCATÁRIA, na forma do Quadro-Resumo, e são devidos por inteiro ainda que o contrato termine antes do prazo mínimo.
13.2. A LOCATÁRIA garantirá acesso adequado ao local de descarga. Custos extras causados por acesso inadequado, espera ou recusa de recebimento serão cobrados dela.
13.3. Os riscos do transporte correm por conta de quem o contratar. Se a LOCATÁRIA contratar ou indicar o transportador, responde pelos danos ao EQUIPAMENTO durante o transporte, a carga e a descarga.

#H CLÁUSULA 14ª. DO COMBUSTÍVEL
14.1. O combustível é fornecido pela LOCATÁRIA. Se a LOCADORA adiantar combustível a pedido da LOCATÁRIA ou de seus prepostos, o valor comprovado será reembolsado na fatura seguinte.

#H CLÁUSULA 15ª. DA RESCISÃO
15.1. A LOCATÁRIA pode encerrar a locação antes do prazo mínimo mediante aviso escrito de 15 (quinze) dias, pagando os aluguéis até a devolução, a desmobilização e multa compensatória de 30% (trinta por cento) sobre os aluguéis mínimos que faltariam até o fim do prazo mínimo.
15.2. A mesma multa compensatória incide na resolução por culpa da LOCATÁRIA (Cláusula 8.3), sem prejuízo da indenização por danos ao EQUIPAMENTO. A multa compensatória é o mínimo da indenização, e a LOCADORA poderá exigir indenização suplementar se provar prejuízo maior (art. 416, parágrafo único, do Código Civil).
15.3. A LOCATÁRIA pode resolver o contrato sem multa se a LOCADORA descumprir a Cláusula 10.4, pagando apenas o período utilizado.
15.4. A falência, a insolvência, o protesto de títulos ou a paralisação geral das obras da LOCATÁRIA autorizam a LOCADORA a resolver o contrato e retomar o EQUIPAMENTO. Em recuperação judicial ou extrajudicial da LOCATÁRIA, os aluguéis vencidos após o pedido serão pagos em dia, e o seu inadimplemento autoriza as medidas da Cláusula 8ª, observada a Lei nº 11.101/2005.
15.5. A paralisação da obra por embargo, por falta de pagamento do contratante da LOCATÁRIA ou de órgão público, ou por decisão da própria LOCATÁRIA, não constitui caso fortuito ou força maior oponível à LOCADORA.

#H CLÁUSULA 16ª. DAS COMUNICAÇÕES, DOS PREPOSTOS E DA PROVA
16.1. São válidas as comunicações feitas aos e-mails e números de WhatsApp indicados no Quadro-Resumo, consideradas recebidas na data do envio, salvo aviso de falha na entrega. A troca de contato só vale após comunicação escrita à outra parte.
16.2. Os prepostos indicados no Quadro-Resumo, e quem mais a LOCATÁRIA indicar por escrito, podem receber e devolver o EQUIPAMENTO, aprovar apontamentos e medições, solicitar horas excedentes, trabalho em fins de semana e combustível, e obrigam a LOCATÁRIA.
16.3. Mensagens, áudios, fotos, vídeos, registros de telemetria e apontamentos trocados entre as partes ou seus prepostos constituem meio de prova, e as partes reconhecem sua autenticidade, salvo prova técnica em contrário.

#H CLÁUSULA 17ª. DA ASSINATURA ELETRÔNICA
17.1. Este contrato pode ser assinado eletronicamente por qualquer modalidade de assinatura eletrônica, inclusive pelo gov.br ou por plataforma de assinatura, que as partes aceitam como válida (art. 10, § 2º, da MP nº 2.200-2/2001 e Lei nº 14.063/2020), dispensadas as testemunhas (art. 784, § 4º, do Código de Processo Civil). Se assinado em papel, levará também a assinatura de duas testemunhas.
17.2. Os signatários declaram ter poderes para obrigar as partes que representam, e a LOCATÁRIA reconhece como seus os atos praticados pelos prepostos indicados no Quadro-Resumo.

#H CLÁUSULA 18ª. DISPOSIÇÕES GERAIS
18.1. A tolerância de uma parte quanto ao descumprimento da outra não importa novação, renúncia ou alteração do contrato.
18.2. A invalidade de uma cláusula não afeta as demais.
18.3. Qualquer alteração deste contrato só vale por escrito, admitida a troca de mensagens eletrônicas entre os representantes ou prepostos das partes.
18.4. Os dados pessoais dos representantes, prepostos e do FIADOR serão tratados para a execução deste contrato, para o exercício regular de direitos e para a proteção do crédito (art. 7º, V, VI e X, da Lei nº 13.709/2018).
18.5. As partes se comprometem a cumprir a Lei nº 12.846/2013 (Lei Anticorrupção) na execução deste contrato.
18.6. Este contrato obriga as partes e seus sucessores a qualquer título.

#H CLÁUSULA 19ª. DO FORO
19.1. Fica eleito o foro da Comarca de Campina Grande/PB, domicílio da LOCADORA e lugar do pagamento, para dirimir as questões oriundas deste contrato (art. 63 do Código de Processo Civil).

E, por estarem de acordo, as partes assinam este contrato.
{{cidade_data}}

#A

#P
#T ANEXO I. TERMO DE ENTREGA E VISTORIA
Contrato nº {{numero}}. Equipamento: {{equip}}, série {{equip_serie}}.
Data e hora da entrega: ____/____/________, ____:____. Local: {{local_obra}}.
Horímetro na entrega: __________ h. Nível de combustível: ______. Fotos e vídeo enviados por WhatsApp em: ____/____/________.
[ ] Caçamba, dentes e unhas: ______________________________
[ ] Rompedor, ponteira, buchas e pinos: ______________________________
[ ] Mangueiras, cilindros e vazamentos: ______________________________
[ ] Esteiras ou pneus: ______________________________
[ ] Cabine, vidros, fechaduras e tampa do tanque: ______________________________
[ ] Baterias (marca e número de série): ______________________________
[ ] Painel, alarmes e iluminação: ______________________________
[ ] Acessórios entregues: {{acessorios}}
Observações: ____________________________________________________________
Recebido pela LOCATÁRIA, por seu preposto: ______________________________ CPF: ________________
Entregue pela LOCADORA: ______________________________

#T ANEXO II. TERMO DE DEVOLUÇÃO
Data e hora da devolução: ____/____/________, ____:____. Horímetro na devolução: __________ h. Nível de combustível: ______.
[ ] Sem avarias além do desgaste natural.
[ ] Avarias constatadas (descrever e fotografar): ____________________________________________________________
Pela LOCADORA: ______________________________ Pela LOCATÁRIA: ______________________________`;

if (typeof module !== 'undefined') module.exports = { CAMPOS: CAMPOS, CONTRATO: CONTRATO };

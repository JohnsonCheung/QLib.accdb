Attribute VB_Name = "MxSampSchmDta"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSampSchmDta."
Property Get SampSchmDta1() As SchmDta: SampSchmDta1 = SchmDta(SampSchm1): End Property
Property Get SampSchmDta2() As SchmDta: SampSchmDta2 = SchmDta(SampSchm2): End Property
Property Get SampSchmDta3() As SchmDta: SampSchmDta3 = SchmDta(SampSchm3): End Property
Property Get SampSchmDta4() As SchmDta: SampSchmDta4 = SchmDta(SampSchm4): End Property


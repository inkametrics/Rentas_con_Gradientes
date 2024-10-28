# RENTAS CON GRADIENTES
---
INKAPITALES
[Inkapitales](https://inkapitales.com/)
---
En colaboración
[Cibertec](https://www.cibertec.org)
---

### Calculadora Rentas con Gradientes

- Se desarrollaron UDFs (Funciones Definidas por el Usuario)

---
Descripción

Complemento de Excel con Funciones personalizadas con Excel VBA, 
diseñado para dar soporte a departamentos crediticios de entidades 
de la industria de finanzas y microfinanzas. 

Asimismo, tiene la potencialidad de ayudar a la comunidad académica
que se encuentra abordando los cursos de (i) matemática 
financiera o (ii) ingeniería económica, de las facultades de 
economía, finanzas e ingeniería.

Se definen por su horizonte, como funciones:
- Gradientes finitas
- Gradientes perpetuas

Y por su patrón de cambio, como:
- Gradientes geométricas
- Gradientes aritméticas

** Las fórmulas y UDFs de este complemento heredan la sintaxis de las 
fórmulas financieras nativas de MS Excel con el objeto que su uso 
minimice posibles fricciones con usuarios financieros recurrentes de
la suit de Microsoft.

## Funciones: Descripción

### Valor Presente
- PV_RGG:         Valor presente de una renta con gradiente geométrica 
- PV_RGG_cPER:    Valor presente de una renta con gradiente geométrica con perpetuidad 
- PV_RGA:         Valor presente de una renta con gradiente aritmética 
- PV_RGA_cPER:    Valor presente de una renta con gradiente aritmética con perpetuidad 

### Valor Futuro
- FV_RGG:         Valor futuro de una renta con gradiente geométrica 
- FV_RGG_cPER:    Valor futuro de una renta con gradiente geométrica con perpetuidad (Infinito / No Existe) 
- FV_RGA:         Valor futuro de una renta con gradiente aritmética 
- FV_RGA_cPER:    Valor futuro de una renta con gradiente aritmética con perpetuidad (Infinito / No Existe)

### Rentas iniciales
- PMT_RGG         Renta inicial de una renta con gradiente geométrica
- PMT_RGG_cPER    Renta inicial de una renta con gradiente geométrica con perpetuidad
- PMT_RGA         Renta inicial de una renta con gradiente aritmética
- PMT_RGA_cPER    Renta inicial de una renta con gradiente aritmética con perpetuidad

### Factores de actualización
- Factor_RU         Factor de actualización de una renta uniforme
- Factor_RGG        Factor de actualización de una renta con gradiente geométrica
- Factor_RGG_cPER   Factor de actualización de una renta con gradiente geométrica con perpetuidad
- Factor2_RGA       Factor de actualización de una renta con gradiente aritmética (factor del segundo monomio)
- Factor2_RGA_cPER  Factor de actualización de una renta con gradiente aritmética con perpetuidad (factor del segundo monomio)

# Funciones: Fórmulas

Todas las fórmulas internamente adoptan la sintaxis básicas de las
funciones básicas de MS Excel. En ese sentido:
- PV:   Present Value / Valor Presente
- R:    Renta uniforme
- R1:   Renta inicial de la serie con gradiente
- t%:   Tasa de descuento o interés, expresado como tasa efectiva en
en mismo periodo que el flujo de caja.
- g%:   Gradiente geométrica o tasa de crecimiento geométrico, expresado como porcentaje
- G:    Gradiente  aritmética, expresado en las mismas unidades monetarias
que la renta.

En adelante se explicitan las fórmulas provenientes de teoría de rentas:

## PV_RGG finita

$$\begin{equation*}
PV_{RGG} = \left\{\begin{matrix}
R_1 \left[\frac{1-(\frac{1+g\%}{1+t\%})^{nper}}{(t\%-g\%)}\right]& ; & t\% <> g\% \\
R_1 \left[\frac{nper}{(1+t\%)}\right] & ; &  t\% = g\%\\
\end{matrix}\right.
\end{equation*}$$

## PV_RGG perpetua

$$\begin{equation*}
PV_{RGG_{Perp}} = \left\{\begin{matrix}
\frac{R_1}{t\% - g\%} & ; & t\% > g\% \\
\infty & ; &  t\% <= g\%\\
\end{matrix}\right.
\end{equation*}$$



## PV_RGA finita

$$PV_{RGA} = R_1 \left[\frac{1-(1+t\%)^{-nper}}{t\%}\right] + 
\frac{G}{t\%} \left[\frac{1-(1+t\%)^{-nper}}{t\%} - \frac{nper}{(1+t\%)^{nper}} \right]$$


## PV_RGA perpetua

$$PV_{RGA_{Perp}} = \left[ \frac{R_1}{{t\%}}\right] + \left[ \frac{G}{{t\%}^2}\right]$$

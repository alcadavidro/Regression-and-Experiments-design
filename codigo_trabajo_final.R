############ Selección del directorio de trabajo --------------
setwd("~/UNAL/Especialización Análitica/Regresión y diseño de experimentos/Trabajo Final")

############ Abriendo las librerias necesarias -------------------
library(openxlsx)               # Libreria para abrir archivos xlxs
library(dplyr)                  # Libreria para manipulación de los datos
library(PerformanceAnalytics)   # Libreria para visualización de los datos
library(ParamHelpers)           # Libreria para apoyo para funciones de optimización
library(mlr)                    # Libreria para tratar datos No disponibles (NAs) 
library(caret)                  # Libreria para definir tamaños de entrenamiento y prueba para modelos
library(gamlss)                 # Libreria para modelos gamlss
library(e1071)                  # Libreria para imputación de datos
library(pastecs)                # Libreria para estadísticos descriptivos
library(mctest)                 # Libreria para evaluar multicolinealidad
library(model)                  # Libreria para evaluar la falta de ajuste

########### Importación y tratamiento de datos ############

# Importamos los datos
datos <- read.xlsx(xlsxFile = "Datos.xlsx",       # Selección del archivo
                   sheet = "Datos",               # Selección de la hoja
                   colNames = T,                  # Lectura de primera fila como nombres
                   rows = 1:384,                  # Lectura de las filas
                   cols = 1:50,                   # Lectura de las columnas
                   detectDates = TRUE)            # Detección de fechas

# Exploración del tipo de datos para su procesamiento 

str(datos)                                        # Visualización de la estructura de los datos

# Primero definimos nuestras variables de interes por el estudio. Por lo que quitamos variables

datos <- subset(datos,
                select = -c(dia, mes, semana))    # Se quitan variables de tiempo

datos <- subset(datos,
                select = -c(expansion_autoclave,
                            barras_14_dias,
                            contenido_aire,
                            falso_fraguado,
                            consistencia_normal,
                            fraguado_final,
                            fraguado_inicial,
                            agua_cemento, flujo,
                            silo_almacenamiento))

datos <- subset(datos,                           # Se quitan las variables de formulación humeda
                select = -c(densidad,
                            clinker_bh,
                            yeso_bh,
                            caliza_bh,
                            escoria_bh,
                            puzolana_bh,
                            ceniza_bh,
                            mat_repro_bh,
                            toneladas,          # Se quita toneladas del lote muestreado
                            analista_fisico))   # Se quita la variable de análista físico

datos <- subset(datos,
                select = -c(alcalis_equivalente,
                            residuo_insoluble))

# Se valida la estructura de los datos 

str(datos)

# Es necesario redefinir las variables 'procedencia_clinker' y 'molino', ya que son categoricas

# Para la variable procedencia del clinker

unique(datos$procedencia_clinker) # 2 niveles

datos$procedencia_clinker <- factor(datos$procedencia_clinker, 
                              levels = c("Internacional", "Local"), 
                              labels = c("Internacional", "Local"))
# Para la variable molino de cemento

unique(datos$molino) # 3 niveles

datos$molino <- factor(datos$molino, 
                       levels = c("Molino 1","Molino 2","Molino 3"),
                       labels = c("Molino 1","Molino 2","Molino 3"))

# Valores faltantes

names(datos[,colSums(is.na(datos)) !=0])      # Validación de valores faltantes

# Dado que son variables cuantitativas, se va imputar la media de la variable en los faltantes

imputed_data <- impute(
    datos,
    classes = list(numeric = imputeMean()))   # Uso de la media

datos <- imputed_data$data                    # Se reemplazan los datos

names(datos[,colSums(is.na(datos)) !=0])      # Se verifica si se imputaron los valores faltantes

# Se verifica si hay falta de registros, o registros con valor cero en las variables de interes

summary(datos)                                # Se verifica si el mínimo es cero, puede indicar

# Para las resistencias, hay registros faltantes como cero. Se pasa a tratar

faltantes.r1 <- which(datos$R1_dia == 0)      # registros no tienen registos de resistencia a un día. 
faltantes.r3 <- which(datos$R3_dias == 0)     # registros no tienen registos de resistencia a 3 días. 
faltantes.r7 <- which(datos$R7_dias == 0)     # registros no tienen registos de resistencia a 7 días.
faltantes.r28 <- which(datos$R28_dias == 0)   # registros no tienen registos de resistencia a 28 días.

length(faltantes.r1)/dim(datos)[1]            # Es el 0.78 %
length(faltantes.r3)/dim(datos)[1]            # Es el 2.3 %
length(faltantes.r7)/dim(datos)[1]            # Es el 2.9 %
length(faltantes.r28)/dim(datos)[1]           # Es el 7.8%

datos[faltantes.r1, "R1_dia"] <- NA
datos[faltantes.r3, "R3_dias"] <- NA
datos[faltantes.r7, "R7_dias"] <- NA
datos[faltantes.r28, "R28_dias"] <- NA

imputed_data <- impute(
    datos,
    classes = list(numeric = imputeMedian()))         # Uso de la media
datos <- imputed_data$data          # Reemplazamos los datos



# Se procede a separar las variables cuantitativas y cualitativas para hacer un análisis exploratorio

cat_var <- dplyr::select(
    .data = datos,
    procedencia_clinker, molino)              # Matriz de variables categoricas

cuant_var <- subset(
    datos, select = -c(procedencia_clinker,
                       molino))               # Matriz de variables cuantitativas

########### Exploración univariada y multivariada de los datos ############

# Para las variables cualitativas

original.values <- par()                      # Valores originales de los parametros
par(bty = 'l')                                # Marco alrededor de los gráficos
par(mfrow = c(1,2))                           # Para configurar dos gráficos e una sola gráfica

# Procedencia del clinker

as.data.frame(x = (sort(prop.table(
    table(cat_var$procedencia_clinker)),
    decreasing = T))) # frecuencias relativas

# el 74% de los datos de procedencia de clinker es local, el 26% es internacional

procedencia <- barplot(table(cat_var$procedencia_clinker),col="gray",
                       ylim = c(0, length(cat_var$procedencia_clinker)),
                       xlab='Procedencia de clinker para molienda',
                       ylab='Frecuencia absoluta',
                       main = "Procedencia",las=1,
                       cex = 0.6, cex.lab=0.8, cex.axis = 0.9)
text(x=procedencia, y=table(cat_var$procedencia_clinker), pos=3, cex=0.8, col="black",
     label=round(table(cat_var$procedencia_clinker), 4))

# Molino usado para la producción de cemento

as.data.frame(x = (sort(prop.table(
    table(cat_variables$molino)),
    decreasing = T))) # frecuencias relativas

# El 82% de los registros provienen del molino 1, el 13% del molino 3 y el 4.7% del molino 2

mol <- unique(datos$molino)

molienda <- barplot(table(cat_var$molino),col="gray", # Datos y color
                    ylim = c(0, length(cat_var$molino)), # Limite del eje y
                    ylab='Frecuencia absoluta', # Label del eje y
                    main = "Molino de cemento",las=1, # Titulo y posición perpendicular de los labels
                    cex = 0.8, cex.lab=0.7,cex.axis = 0.9, # Tamaño de la letra
                    xaxt="n") # Quita los labels del eje x
text(x=mol, y=par()$usr[3]-0.1*(par()$usr[4]-par()$usr[3]),
     labels=mol, srt=45, adj=1, xpd=TRUE, cex = 0.8) # Añade el label del eje x rotado en 45°
text(x=molienda, y=table(cat_var$molino), pos=3, cex=0.8, col="black",
     label=round(table(cat_var$molino), 4)) # Añade los datos a las barras

# Para variables cuantitativas

# resumen de los estadísticos descriptivos

options(scipen = 100) # Penalidad para indicar impresión de números en su notación real
options(digits = 2) # Máximo de dígitos
stat.desc(cuant_var) 

# La variable respuesta de interes, es el desarrollo de resistencias a la compresión, se verifica la distribución de estás variables

Resist <- dplyr::select(datos,                 # Matriz de datos de resistencias
                        R1_dia,R3_dias,        # Resistencias a 1 día y 3 días
                        R7_dias, R28_dias)     # Resistencias a 7 y 28 días

apply(Resist, 2, function(x)kurtosis(x))       # Kurtosis de las resistencias
                                               # La de un día y 3 días, son leptocurticas
                                               # Las de 7 y 28 días tienden a mesocurtica

apply(Resist, 2, function(x)skewness(x))       # Sesgo de las resistencias, todas tienen u sesgo a la derecha

# Corroboración visual
par(mfrow = c(1,1))
boxplot(Resist, main = "Resistencias a la diferentes edades", 
        ylab = "Psi")                          # Se evidencia los valores obtenidos anteriormente

# Por conocimiento previo de expertos, se sabe que hay correlaciones entre algunas variables
# Pero no es tan simple hacer una detección de la misma con métodos como regresión lineal o 
# Matriz de correlación, por lo que se opta por usar métodos como el VIF, Tolerance Limit, CN, CI.

matriz_cor <- cor(cuant_var)                   # Matriz de correlaciones


mod_full <- lm(R1_dia ~. - R3_dias - R7_dias - R28_dias, data = datos)
summary(mod_full)                              # Dada algunas varianzas de los regresores
                                               # y la importancia de los mismos, se sabe que se presenta multicolinealidad

# Se sacan las variables respuestas de la matriz de variables cuantitativas

cuant_regressors <- subset(cuant_var, select = -c(R1_dia, R3_dias, R7_dias, R28_dias))

stand_regressor <- scale(cuant_regressors)

omcdiag(x = stand_regressor,
        y = Resist$R1_dia, Inter = T)          # Hay alta evidencia de colinealidad
omcdiag(x = cuant_regressors,
        y = Resist$R3_dias, Inter = T)          # Hay alta evidencia de colinealidad
omcdiag(x = cuant_regressors,
        y = Resist$R7_dias, Inter = T)          # Hay alta evidencia de colinealidad
omcdiag(x = cuant_regressors,
        y = Resist$R28_dias, Inter = T)          # Hay alta evidencia de colinealidad

imcdiag(x = cuant_regressors,
        y = Resist$R1_dia, all = T)
imcdiag(x = cuant_regressors,
        y = Resist$R3_dias, all = T)
imcdiag(x = cuant_regressors,
        y = Resist$R7_dias, all = T)
imcdiag(x = cuant_regressors,
        y = Resist$R28_dias, all = T)

mc.plot(x = stand_regressor,
        y = Resist$R7_dias)

# Dado la alta existencia de multicolinealidad entre los regresores, se pueden tomar múltiples caminos 
# se opta por remover aquellos que muestran los valores propios más bajos y a consideración del experto para ir quitando
# información redundante. 

cuant_regressors <- subset(cuant_regressors,
                           select = -c(mat_repro_bs,
                                       ceniza_bs,
                                       puzolana_bs,
                                       escoria_bs,
                                       caliza_bs,
                                       oxido_sodio,
                                       oxido_potasio,
                                       oxido_titanio,
                                       perdida_ignea))

# Se verifica una vez más la multicolinealidad general e individual. 

stand_regressor <- scale(cuant_regressors)
omcdiag(x = cuant_regressors,
        y = Resist$R1_dia)
imcdiag(x = cuant_regressors,
        y = Resist$R1_dia, all = T)

mc.plot(x = stand_regressor,
        y = Resist$R1_dia)

# Se sigue manteniendo alta colinealidad entre algunos regresores pero, por consejo de expertos, es 
# mejor optar por un backward selection con un modelo que se vea mejor 

datos_modelo <- cbind(cuant_regressors,cat_var, Resist) # Creación de una matriz de datos de ínteres

# Se inicia la comprobación de diferentes modelos, por decisión del experto se define la variable respuesta 
# la resistencia a 7 días. 

datos_modelo <- subset(datos_modelo, 
                       select = -c(R3_dias, R28_dias, R1_dia))

################# Modelos de regresión ############

panel.cor <- function(x, y, digits = 2, prefix = "", cex.cor, ...)
{
    usr <- par("usr"); on.exit(par(usr))
    par(usr = c(0, 1, 0, 1))
    r <- cor(x, y)
    txt <- format(c(r, 0.123456789), digits = digits)[1]
    txt <- paste(prefix, txt, sep = "")
    if(missing(cex.cor)) cex.cor <- 0.8/strwidth(txt)
    text(0.5, 0.5, txt, cex = cex.cor * r)
}

pairs(cuant_regressors, 
      pch=19, las=1,
      upper.panel = panel.smooth, lower.panel = panel.cor)
corrr <- cor(cuant_regressors)

### Modelo 1 

mod1 <- lm(R7_dias ~., data = datos_modelo)   # Modelo 1, modelo de regresión lineal con todas las variables

summary(mod1)
par(mfrow = c(2,2))
plot(mod1)

# Se producen outliers, es necesario sacarlos

hii_mod1 <- lm.influence(mod1)$hat # hii del modelo 1
which.max(hii_mod1) # El punto 189 tiene alta influencia. 
eii_mod1 <- mod1$residuals # residuales del modelo 1

dii_mod1 <- (eii_mod1-mean(eii_mod1))/summary(mod1)$sigma  # Residuales estandarizados

y_ajus_mod1 <- mod1$fitted.values # Valores ajustados del modelo 

ri_mod1 <- eii_mod1/sqrt(summary(mod1)$sigma^2*(1-hii_mod1)) # Residuales estudentizados

PRESS_mod1 <- eii_mod1/(1-hii_mod1)

datos_modelo1 <- cbind(datos_modelo, hii = hii_mod1,
                       eii = eii_mod1, dii = dii_mod1, y_ajustado = y_ajus_mod1,
                       rii = ri_mod1, PRESS = PRESS_mod1)

sort(hii_mod1, decreasing = T)

leverage_mod1 <- which(hii_mod1 > 0.15) # Puntos de alta influencia

cutoff <- 4/((nrow(mod2$model)-length(mod1$coefficients)-2)) # Corte de influencia 
plot(mod2, which=4, cook.levels=cutoff) # Gráfico de distancia de cook

### Modelo 1 sin los puntos de influencia

mod1v2 <- lm(R7_dias ~., data = datos_modelo[-leverage_mod1,])

summary(mod1v2)
AIC(mod1)
AIC(mod1v2)
result_mod1v2 <- lack.fit(mod1v2)
result_mod1v2

# Se logró bajar el AIC de 5629 a 5378 quitando los puntos de mayor influencia

## Modelo 2 
datos_modelo1v2 <- datos_modelo[-leverage_mod1,]
datos_modelo1v2
mod2 <- stepAIC(mod1v2,direction = "backward", trace = F, k = log(dim(datos_modelo1v2)[2]))
summary(mod2)
par(mfrow = c(2,2))
plot(mod2)


################
#hii_mod2 <- lm.influence(mod2)$hat # hii del modelo 1
#leverage_mod2 <- which(hii_mod2 > 0.1) # Puntos de alta influencia
#datos_modelo2v2 <- datos_modelo1v2[-leverage_mod2,]


#mod1v2 <- lm(R7_dias ~., data = datos_modelo2v2)
#mod2v2 <- stepAIC(mod1v2,direction = "backward", trace = F, k = log(dim(datos_modelo1v2)[2]))
#summary(mod2v2)


#plot(mod2v2)

#AIC(mod2v2)

###### Modelo 3 #####

stepGAICAll.A(object, scope = NULL, sigma.scope = NULL, nu.scope = NULL, tau.scope = NULL)

mod3.empty <- gamlss(formula = R7_dias ~ 1,
                     sigma.formula = ~ 1,
                     family = NO,
                     data = datos_modelo1v2,
                     control=gamlss.control(n.cyc=500))

summary(mod3.empty)
formula.mod <- formula( ~ blaine + retenido_325 + oxido_silice + oxido_alumina + 
                          oxido_hierro + oxido_calcio + oxido_magnesio + trioxido_azufre + 
                           clinker_bs + yeso_bs + procedencia_clinker + molino)
mod3 <- stepGAICAll.A(object=mod3.empty, k=2, 
                      scope=list(lower= ~ 1, upper= formula.mod), 
                      sigma.scope=list(lower= ~ 1, upper= formula.mod))
options()
summary(mod3)
plot(mod3)
wp(mod3)

mod3.lm <- lm(R7_dias ~ clinker_bs + molino, data = datos_modelo1v2) 
summary(mod3.lm)

fitR7 <- fitDist(datos_modelo1v2$R7_dias, type = "realplus")
fitR7$fits # La distribución que mejor se ajusta es la Box - Cox t

mod3v2.empty <- gamlss(formula = R7_dias ~ 1,
                       sigma.formula = ~1,
                       nu.formula = ~1,
                       tau.formula = ~1,
                       data = datos_modelo1v2,
                       control=gamlss.control(n.cyc=500),
                       family =BCTo)
mod3v2 <- stepGAICAll.A(object=mod3v2.empty, k=2, 
                      scope=list(lower= ~ 1, upper= formula.mod), 
                      sigma.scope=list(lower= ~ 1, upper= formula.mod),
                      tau.scope = list(lower= ~ 1, upper = formula.mod),
                      nu.scope = list(lower = ~1, upper = formula.mod))
summary(mod3v2)
plot(mod3v2)
wp(mod3v2)

mod3v3.empty <- gamlss(formula = R7_dias ~ 1,
                       sigma.formula = ~1,
                       data = datos_modelo1v2,
                       control=gamlss.control(n.cyc=500),
                       family =GA)
mod3v3 <- stepGAICAll.A(object=mod3v3.empty, k=2, 
                        scope=list(lower= ~ 1, upper= formula.mod), 
                        sigma.scope=list(lower= ~ 1, upper= formula.mod))
summary(mod3v3)
plot(mod3v3)
wp(mod3v3)



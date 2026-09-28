# Pruebas de la skill `formatear-curs-clinic-argos`

Simulador mínimo de la API de Word (`mock-word.js`) y un Curs Clínic de prueba en sucio
(`fixture.js`) con todos los casos: notas de MI/MFiC/INFERMERIA con barras "CC" y líneas grises,
nota antigua v1, PROVES en formato antiguo y nuevo (en castellano, con recuento de grupo, estado
"Anulada", determinaciones que no están en el diccionario, una imagen entre pruebas),
INTERCONSULTES pegadas de la app con "veure CC", "No realizadas (0)" y notas del curs clínic.

```
node run.js ../../skills/formatear-curs-clinic-argos/SKILL.md                 # pasada + 2ª pasada (debe quedar igual)
node run.js ../../skills/formatear-curs-clinic-argos/SKILL.md --incremental   # + pegado del día siguiente
node run.js ../../skills/formatear-curs-clinic-argos/SKILL.md --api14         # Word sin WordApi 1.5
node run.js ../../skills/formatear-curs-clinic-argos/SKILL.md --sin-plantilla # sin estilos MiniEspacio/Lista con viñetas
node run.js <SKILL.md> --volcar salida.txt                                    # guarda el documento resultante para compararlo
```

El simulador no es Word: sirve para comprobar la lógica y comparar versiones de la skill entre sí,
no sustituye a una primera prueba sobre una copia de un Curs Clínic real.

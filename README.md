# Excel-Tools
Utilidades para uso en excel

## Separar Texto
Formula para separar texto de una cadena separada por algun caracter.
por Ejemplo,en la siguiente cadena, se podría separar cada uno de los elementos en columnas diferentes de la siguente manera
<table>
  <thead>
    <tr>
      <th>Cadena con caracteres</th>
      <th>PO</th>
      <th>Estilo</th>
      <th>Talla</th>
      <th>Cantidad</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>PO0034924-2$C10P-C01$2XL$100</td>
      <td>PO0034924-2</td>
      <td>C10P-C01</td>
      <td>2XL</td>
      <td>100</td>
    </tr>
  </tbody>
</table>

* PO: ``` =SepararTexto(A2,1,"$")```  
* Estilo: ``` =SepararTexto(A2,2,"$")```
* Talla: ``` =SepararTexto(A2,3,"$")```
* Cantidad: ``` =SepararTexto(A2,4,"$")```

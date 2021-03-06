# xlMdDna
## Excel mermaid-String to Diagram

### Usage
1. [xll 32bit※](https://github.com/WaToI/xlMdDna/blob/master/xlMdDna-AddIn-packed.xll) download&open xll in Excel
- cell A1:  graph LR   
- cell A2:  HelloWorld  
- cell c3: =mermaid(A1:A2)
- autoOpen Live-Preview-window
※ if yourExcel is 64bit [xll 64bit](https://github.com/WaToI/xlMdDna/blob/master/xlMdDna-AddIn64-packed.xll)


### mermaid
```
graph TD  
 　A-->B  
 　A-->C  
 　B-->D  
 　C-->D  
```

```
sequenceDiagram
    participant Alice
    participant Bob
    Alice->>John: Hello John, how are you?
    loop Healthcheck
        John->>John: Fight against hypochondria
    end
    Note right of John: Rational thoughts <br/>prevail...
    John-->>Alice: Great!
    John->>Bob: How about you?
    Bob-->>John: Jolly good!
```
### NotSupport
- classDiagram 
- gantt 

### Credits
**thanks**
* [Excel-DNA](https://excel-dna.net/)  
* [mermaid](http://knsv.github.io/mermaid/#downstream-projects)

---
License MIT
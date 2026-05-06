# PeriTAB

Suplemento para Microsoft Word desenvolvido para automatizar e agilizar a elaboração de laudos periciais.

<p align="center">
  <!-- <img src="PeriTAB/Resources/peritab-readme.png" width="600"/> -->
</p>

---

## 📌 Visão Geral

O **PeriTAB** é um add-in para Microsoft Word que adiciona uma aba personalizada contendo ferramentas voltadas à elaboração de laudos periciais.

O objetivo é aumentar a produtividade, padronizar documentos e reduzir erros operacionais durante a confecção de laudos.

O projeto combina:

- Integração com Microsoft Word (Add-in)
- Interface em forma de aba personalizada (Ribbon)
- Funcionalidades automatizadas em C#
- Ferramentas voltadas à produção de laudos

---

## 🧩 Funcionalidades

A aba **PeriTAB** adicionada ao Word oferece:

- Inserção automatizada de textos e estruturas padrão
- Execução de rotinas para agilizar tarefas repetitivas
- Padronização de elementos do laudo
- Facilidades para edição e organização do documento

---

## 🧱 Arquitetura

O sistema é baseado em:

- **C#** para o desenvolvimento do add-in
- **VSTO / COM Add-in** para integração com o Microsoft Word
- Integração direta com o modelo de objetos do Word (Interop)

---

## ⚙️ Funcionamento

Fluxo típico de uso:

1. Usuário abre o Microsoft Word
2. A aba **PeriTAB** é carregada automaticamente
3. Usuário utiliza os botões disponíveis na aba
4. As rotinas em C# são executadas

---

### Pré-requisitos

- Microsoft Word
- Windows

---

### Instalação

1. Baixe a versão mais recente na aba **Releases**
2. Execute o instalador (se disponível)
3. Abra o Microsoft Word
4. A aba **PeriTAB** estará disponível na interface

---

## 🖥️ Compatibilidade

Testado apenas no:

- Microsoft Word 16
- Windows 10 e 11

---

## 📌 Versionamento

O projeto segue versionamento semântico:

- `v0.x` → versões em desenvolvimento
- `v1.0.0` → primeira versão estável
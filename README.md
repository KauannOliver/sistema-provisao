# Sistema de Lançamento de Provisões e Estornos

Este projeto foi desenvolvido para gerenciar o processo de lançamento, acompanhamento e estorno de provisões, facilitando a análise de quais provisões já foram estornadas, parcial ou totalmente, além de possibilitar o cadastro de novos clientes e dados financeiros. Toda a base de dados é armazenada e manipulada diretamente em planilhas Excel, proporcionando praticidade para importação e exportação de dados.

FUNCIONALIDADES PRINCIPAIS

1. CRUD Completo de Provisões, Estornos e Clientes: O sistema oferece funcionalidades de Create, Read, Update e Delete (CRUD) para gerenciar provisões, estornos e clientes. Isso permite um controle completo sobre os dados financeiros lançados no sistema, incluindo a possibilidade de edição e exclusão de informações.

2. Importação de Dados a Partir de Excel: Não é necessário cadastrar provisões manualmente uma por uma. O sistema permite a importação de dados diretamente de um modelo Excel próprio, facilitando o processo de subida de informações em massa para a base de dados.

3. Acompanhamento de Provisões e Estornos: O sistema realiza o acompanhamento detalhado de provisões, permitindo verificar quais já foram totalmente estornadas, parcialmente estornadas ou ainda estão pendentes de estorno. A visualização é clara e permite fácil controle sobre o que já foi estornado e o que ainda está pendente.

4. Exportação de Planilhas de Provisões Pendentes: É possível exportar uma planilha que contém apenas as provisões pendentes, juntamente com informações detalhadas sobre quanto já foi estornado e quanto ainda falta estornar. Isso ajuda no controle das obrigações financeiras e facilita a gestão dos valores que precisam ser ajustados.

5. Manipulação e Formatação de Dados Financeiros: O sistema utiliza formatação contábil para todos os valores financeiros, como receita bruta, impostos (ICMS, ISS, PIS, COFINS, CPRB), e valores estornados, garantindo que todos os dados estejam corretos e bem organizados para análise.

6. Relatórios e Controle: O sistema permite a geração de relatórios detalhados para acompanhamento das provisões e estornos, oferecendo uma visão clara e organizada dos valores e status de cada lançamento.

TECNOLOGIAS UTILIZADAS

1. Flet: Framework utilizado para o desenvolvimento do front-end, proporcionando uma interface moderna e fácil de usar.

2. Python: Utilizado para o back-end do sistema, implementando a lógica de negócio e integração com Excel.

3. Pandas & Openpyxl: Bibliotecas essenciais para manipulação de dados e interação com planilhas Excel.

4. Excel: Utilizado como banco de dados principal para armazenar provisões, estornos e clientes, garantindo uma fácil integração e portabilidade dos dados.


CONCLUSÃO

Este projeto foi desenvolvido para proporcionar um sistema completo de gestão de provisões e estornos, com funcionalidades que otimizam o controle financeiro de empresas. Com a possibilidade de importar e exportar dados em Excel, o sistema oferece flexibilidade para quem já utiliza planilhas como base de dados. Além disso, o CRUD completo, o acompanhamento detalhado de estornos e a exportação de provisões pendentes tornam este sistema uma ferramenta poderosa para a gestão financeira.

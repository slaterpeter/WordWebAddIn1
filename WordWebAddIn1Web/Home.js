
(function () {
    "use strict";

    var messageBanner;

    // A função inicializar deverá ser executada cada vez que uma nova página for carregada.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Inicializar o mecanismo de notificação e ocultá-lo
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // Se não estiver usando o Word 2016, use a lógica de fallback.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("Este exemplo exibe o texto selecionado.");
                $('#button-text').text("Exibir!");
                $('#button-desc').text("Exibir o texto selecionado");
                
                $('#highlight-button').click(displaySelectedText);
                return;
            }

            $("#template-description").text("Este exemplo realça a palavra mais longa no texto que você selecionou no documento.");
            $('#button-text').text("Realçar!");
            $('#button-desc').text("Realça a palavra mais longa.");
            
            loadSampleData();

            // Adicione um manipulador de eventos de clique ao botão de realce.
            $('#highlight-button').click(hightlightLongestWord);
        });
    };

    function loadSampleData() {
        // Execute uma operação em lote com base no modelo de objeto do Word.
        Word.run(function (context) {
            // Crie um objeto de proxy para o corpo do documento.
            var body = context.document.body;

            // Coloque um comando na fila para limpar o conteúdo do corpo.
            body.clear();
            // Coloque um comando na fila para inserir texto no final do corpo do documento do Word.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Sincronize o estado do documento executando os comandos na fila e retorne uma promessa para indicar a conclusão da tarefa.
            return context.sync();
        })
        .catch(errorHandler);
    }

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Coloque um comando na fila para obter a seleção atual e, em seguida,
            // crie um objeto de intervalo de proxy com os resultados.
            var range = context.document.getSelection();
            
            // Essa variável manterá os resultados da pesquisa para a palavra mais longa.
            var searchResults;
            
            // Coloque um comando na fila para carregar o resultado da seleção de intervalo.
            context.load(range, 'text');

            // Sincronize o estado do documento executando os comandos na fila
            // e retorne uma promessa para indicar a conclusão da tarefa.
            return context.sync()
                .then(function () {
                    // Obtenha a palavra mais longa da seleção.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Coloque um comando de pesquisa na fila.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Coloque um comando na fila para carregar a propriedade de fonte dos resultados.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Coloque um comando na fila para realçar os resultados da pesquisa.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Amarelo
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
        .catch(errorHandler);
    } 


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('O texto selecionado é:', '"' + result.value + '"');
                } else {
                    showNotification('Erro:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Erro:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Função auxiliar para exibir notificações
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();

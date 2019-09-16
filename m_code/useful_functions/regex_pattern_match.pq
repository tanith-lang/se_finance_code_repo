(string) => 
let
    fx = Web.Page(
            "<script>
                var x = '"&string&"';
                var y = new RegExp(/^((\w+)\-(\d+)\-(\d+))$/g);
                var z = y.test(x);
                document.write(z);
            </script>"
    ),
    output = Logical.From(fx[Data]{0}[Children]{0}[Children]{1}[Text]{0})
in
    output
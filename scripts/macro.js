async function pasteTypstDocx() {
    try {
        const result = await (await fetch('http://127.0.0.1:5180')).text();
        if (result.toLowerCase().startsWith('error')) {
            alert(result);
        } else if (result !== '') {
            Selection.InsertFile(
                result,
                undefined,
                undefined,
                undefined,
                undefined
            );
        }
    } catch (e) {
        alert(e);
    }
}

async function insertTypstDocx() {
    Selection.Copy();
    const path = await (await fetch('http://127.0.0.1:5180')).text();
    if (path !== '') {
        Selection.Delete();
        Selection.InsertFile(path, undefined, undefined, undefined, undefined);
    }
}

const { converter, getStyle} = require('./app');
const target = require('./model.json');

let source = 'Stage1.xlsx'
let output = 'converted.json'

describe('Spreadsheet converter', () => {

    it('converter function exists', () => {
        expect(converter).toBeDefined();
    });

    test('Returns a json object', async () => {
        const check = await converter(source, output);
        expect(typeof check).toBe('object');
    });

    test('Returns the correct keys', async () => {
        const check = await converter(source, output);
        expect(Object.keys(check)).toEqual(Object.keys(target));
    });

    test('Returns the correct styles', async () => {
        const check = await converter(source, output);

        check.styles.forEach((style, index) => {
            target.styles.forEach((targetStyle, targetIndex) => {
                if (style === targetStyle) {
                    expect(style).toEqual(targetStyle);
                }
            })
        })
    })

    test('Returns the correct rows length', async () => {
        const check = await converter(source, output);
        expect(check.rows.length).toEqual(target.rows.length);
    })

    test('Returns the correct columns length', async () => {
        const check = await converter(source, output);
        expect(check.cols.length).toEqual(target.cols.length);
    })

    test('Returns the correct cells keys', async () => {
        const check = await converter(source, output);

        Object.values(check.rows).forEach((row, indexRow) => {
            const cells = Object.values(row.cells);
            const targetKeys = Object.keys(target.rows[indexRow].cells);

            targetKeys.forEach((key, indexKey) => {
                expect(Object.keys(cells[key])).toEqual(Object.keys(target.rows[indexRow].cells[key]));
            })
        })
    })

    test('Returns the correct cells text', async () => {
        const check = await converter(source, output);

        Object.values(check.rows).forEach((row, indexRow) => {
            const cells = Object.values(row.cells);
            const targetKeys = Object.keys(target.rows[indexRow].cells);
            
            targetKeys.forEach((key, indexKey) => {
                expect(cells[indexKey].text).toEqual(target.rows[indexRow].cells[key].text);
            })
        })
    })

    test('Returns the correct cells style', async () => {
        const check = await converter(source, output);

        Object.values(check.rows).forEach((row, indexRow) => {
            const cells = Object.values(row.cells);
            const targetKeys = Object.keys(target.rows[indexRow].cells);
            
            targetKeys.forEach((key, indexKey) => {
                const targetStyleId = target.rows[indexRow].cells[key].style
                const targetStyle = target.styles[targetStyleId];

                const sourceStyleId = cells[key].style;
                const sourceStyle = check.styles[sourceStyleId];

                expect(sourceStyle).toEqual(targetStyle);
            })
        })
    })
    
})
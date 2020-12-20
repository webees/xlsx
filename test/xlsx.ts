
import test from 'ava'
import XLSX from '../src/index'

test('xlsx', async t => {
  await XLSX.start({
    src: 'test/xlsx',
    dist: 'dist',
    conf: '0.xlsx'
  })
  t.pass()
})

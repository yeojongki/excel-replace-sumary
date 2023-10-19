'use client'

import { Form, Button, Upload, Card, Input, UploadProps } from 'antd'
import { UploadOutlined } from '@ant-design/icons'
import IgnoreText from './IgnoreText'
import { type WorkBook } from 'xlsx'
import { RcFile } from 'antd/es/upload'
import {
  EXCEL_ACCEPT,
  exportExcelFile,
  importExcelFromBuffer,
  sheet2JSON,
  splitString,
} from './utils/excel'

interface ReplaceData {
  map: Record<string, string>
  sortedMapData: {
    key: string
    value: string
  }[]
}

export default function Home() {
  const [form] = Form.useForm()

  const parseOriginalSheet = (
    originalExcelData: WorkBook['Sheets'],
    originalSheetName: string,
    originalLocaleColumnNameOriginal: string,
    originalLocaleColumnNameTarget: string,
  ) => {
    const originalSheet = originalExcelData[originalSheetName]
    if (!originalSheet) {
      throw new Error('sheet name not found:' + originalSheet)
    }
    const originalSheetData: Record<string, string>[] =
      sheet2JSON(originalSheet)
    console.log(originalSheetData)

    const result: ReplaceData = {
      map: originalSheetData.reduce((acc, next) => {
        acc[next[originalLocaleColumnNameOriginal]] =
          next[originalLocaleColumnNameTarget]
        return acc
      }, {} as Record<string, string>),
      sortedMapData: originalSheetData
        .map((item) => ({
          key: item[originalLocaleColumnNameOriginal],
          value: item[originalLocaleColumnNameTarget],
        }))
        .sort((a, b) => b.key?.length - a.key?.length)
    }

    return result
  }

  const replaceExcel = (
    replaceSheetData: WorkBook['Sheets'],
    replacedSheetName: string,
    replaceMap: ReplaceData,
    ignoreValues: string[] = [],
  ) => {
    const { map, sortedMapData } = replaceMap
    const sheet = replaceSheetData[replacedSheetName]
    const sheetData = sheet2JSON<string[]>(sheet, 1) // header 1
    const shoudFinalHandle: {
      cell: string
      cellIndex: number
      index: number
    }[] = []

    const newData = sheetData.map((item, index) => {
      return item.map((cell, cellIndex) => {
        if (map[cell]) {
          return map[cell]
        } else {
          // multiple selects
          if (
            !map[cell] &&
            typeof cell === 'string' &&
            cell.indexOf(',') > -1
          ) {
            const splited = splitString(cell)
            let splitMatchCount = 0

            const splitedAndJoin = splited
              .map((splitCell) => {
                const mapSplitVal = sortedMapData.find(
                  (item) => item.key === splitCell,
                )?.value
                if (mapSplitVal) {
                  splitMatchCount += 1
                  return mapSplitVal
                }

                return splitCell
              })
              .join(', ')
            if (splitMatchCount < splited.length) {
              shoudFinalHandle.push({
                cell: splitedAndJoin,
                index,
                cellIndex,
              })
            }

            return splitedAndJoin
          }
        }

        return cell
      })
    })

    const ignoreValuesMap = ignoreValues.reduce((acc, next) => {
      acc[next] = true
      return acc
    }, {} as Record<string, boolean>)

    shoudFinalHandle.forEach((item) => {
      for (const mapItem of sortedMapData) {
        const value = map[mapItem.key]
        if (ignoreValuesMap[value]) {
          continue
        }

        if (item.cell.indexOf(mapItem.key) > -1) {
          const oldCell = newData[item.index][item.cellIndex]
          const newCell = oldCell.replace(mapItem.key, value)
          const isChanged = oldCell !== newCell
          if (isChanged) {
            newData[item.index][item.cellIndex] = newCell
          }
        }
      }
    })

    return newData
  }

  const onFinish = (values: any) => {
    const replaceMap = parseOriginalSheet(
      values.originalExcel[0].response,
      values.originalSheetName,
      values.originalLocaleColumnNameOriginal,
      values.originalLocaleColumnNameTarget,
    )
    const replaced = replaceExcel(
      values.replacedExcel[0].response,
      values.replacedSheetName,
      replaceMap,
      values.replacedIgnoreNames,
    )
    exportExcelFile(
      replaced,
      values.replacedSheetName,
      `new-${values.replacedSheetName}-${+new Date()}.xlsx`,
    )
  }

  const initialValues = {
    originalSheetName: 'd5',
    originalLocaleColumnNameOriginal: '印尼',
    originalLocaleColumnNameTarget: '中文',
    replacedSheetName: 'ID',
  }

  const formItemLayout = {
    labelCol: {
      xs: { span: 24 },
      sm: { span: 24 },
      md: { span: 13 },
    },
    wrapperCol: {
      xs: { span: 24 },
      sm: { span: 24 },
      md: { span: 11 },
    },
  }

  const normFile = (e: any) => {
    return e?.fileList
  }

  const parseOriginalExcel: UploadProps['customRequest'] = async ({
    file,
    onSuccess,
    onError,
  }) => {
    try {
      const excelData = importExcelFromBuffer(
        await (file as RcFile).arrayBuffer(),
      )
      onSuccess?.(excelData)
    } catch (error) {
      onError?.(error as any)
      console.error(error)
    }
  }

  const parseReplaceExcel: UploadProps['customRequest'] = async ({
    file,
    onSuccess,
    onError,
  }) => {
    try {
      const excelData = importExcelFromBuffer(
        await (file as RcFile).arrayBuffer(),
      )
      onSuccess?.(excelData)
    } catch (error) {
      onError?.(error as any)
      console.error(error)
    }
  }

  return (
    <main className='flex min-h-screen flex-col items-center justify-between p-12'>
      <Form
        form={form}
        className='xs:w-full sm:w-[550px] md:w-[700px]'
        onFinish={onFinish}
        initialValues={initialValues}
        {...formItemLayout}
      >
        <Card className='!mb-4' size='small' title='Original excel'>
          <Form.Item
            label='original excel'
            name={'originalExcel'}
            valuePropName='fileList'
            getValueFromEvent={normFile}
            rules={[{ required: true, message: 'Please upload!' }]}
          >
            <Upload accept={EXCEL_ACCEPT} customRequest={parseOriginalExcel}>
              <Button icon={<UploadOutlined />}>Click to Upload</Button>
            </Upload>
          </Form.Item>

          <Form.Item
            label='sheet name'
            name={'originalSheetName'}
            rules={[{ required: true, message: 'Please input!' }]}
          >
            <Input></Input>
          </Form.Item>

          <Form.Item
            label='locale column name original'
            name={'originalLocaleColumnNameOriginal'}
            rules={[{ required: true, message: 'Please input!' }]}
          >
            <Input></Input>
          </Form.Item>

          <Form.Item
            label='locale column target'
            name={'originalLocaleColumnNameTarget'}
            rules={[{ required: true, message: 'Please input!' }]}
          >
            <Input></Input>
          </Form.Item>
        </Card>

        <Card className='!mb-4' size='small' title='Excel to be replaced'>
          <Form.Item
            label='excel to be replaced'
            name={'replacedExcel'}
            rules={[{ required: true, message: 'Please upload!' }]}
            valuePropName='fileList'
            getValueFromEvent={normFile}
          >
            <Upload accept={EXCEL_ACCEPT} customRequest={parseReplaceExcel}>
              <Button icon={<UploadOutlined />}>Click to Upload</Button>
            </Upload>
          </Form.Item>

          <Form.Item
            label='sheet name'
            name={'replacedSheetName'}
            rules={[{ required: true, message: 'Please input!' }]}
          >
            <Input></Input>
          </Form.Item>
        </Card>

        <Card className='!mb-4' size='small' title='Other config'>
          <Form.Item
            className='!mb-0'
            label='ignore text'
            name={'replacedIgnoreNames'}
          >
            <IgnoreText></IgnoreText>
          </Form.Item>
        </Card>

        <div className='text-center'>
          <Button type='primary' htmlType='submit'>
            Submit
          </Button>
          <Button className='ml-3' danger htmlType='reset'>
            Reset
          </Button>
        </div>
      </Form>
    </main>
  )
}

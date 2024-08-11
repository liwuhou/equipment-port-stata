import { useState } from 'react'
import { message, Flex } from 'antd'
import UploadBtn from './components/UploadBtn'
import Result, { ResultStatus } from './components/Result'
import type { WorkBook, WorkSheet } from 'xlsx'

type OriginData = Map<string, MachinRoom>
type HandleData = Map<string, MachinRoom>

interface AnalyzedData {
  originData: OriginData
  handleData: HandleData
  describtipn?: string
}

type AnalyzeWorkBook = (w: WorkBook) => AnalyzedData
type HandleUploadFile = (w: WorkBook) => void
type HandleCheckFile = (w: WorkBook) => boolean

enum SHEET_NAME {
  handle = '原端口',
  origin = '导出表',
}

// 分光器端口状态
enum PortStatus {
  ON,
  OFF,
}

// 分光器端口
interface Port {
  name: string
  status: PortStatus
  next: string
  belongTo: Splitter
}

// 分光器
interface Splitter {
  name: string
  belongTo: MachinRoom
  ports: Map<string, Port>
}

// 机房
interface MachinRoom {
  name: string
  splitters: Map<string, Splitter>
}

const formatSheetData = (s: WorkSheet): Map<string, MachinRoom> => {
  const map = new Map<string, MachinRoom>()
  const maxLength = parseInt(s['!ref']?.split(':')?.[1]?.replace?.(/[a-z]/gi, '') ?? '0', 10)

  if (maxLength <= 2) return map

  for (let i = 3; i <= maxLength; i++) {
    const cell = `C${i}`
    const roomName = Reflect.get(s, cell)?.v ?? ''

    if (roomName) {
      if (!map.has(roomName)) {
        map.set(roomName, { name: roomName, splitters: new Map() as MachinRoom['splitters'] })
      }
      const machinRoom = map.get(roomName)!
      const { splitters } = machinRoom

      const splitterName = Reflect.get(s, `E${i}`)?.v ?? ''

      if (splitterName) {
        if (!splitters.has(splitterName)) {
          splitters.set(splitterName, {
            name: splitterName,
            belongTo: machinRoom,
            ports: new Map() as Splitter['ports'],
          })
        }
        const splitter = splitters.get(splitterName)!
        const { ports } = splitter

        const portName = Reflect.get(s, `F${i}`)?.v ?? ''

        if (portName) {
          if (!ports.has(portName)) {
            const next = Reflect.get(s, `H${i}`)?.v || ''
            const status = Reflect.get(s, `G${i}`)?.v === '在用' ? PortStatus.ON : PortStatus.OFF
            ports.set(portName, { name: portName, belongTo: splitter, status, next })
          }
        }
      }
    }
  }

  return map
}

const diffEquipmentData = (originData: Map<string, MachinRoom>, o: Map<string, MachinRoom>): JSX.Element[] => {
  return Array.from(originData.values()).map((h) => {
    const machiRoomName = h.name
    const splitterCount = h.splitters.size
    let [onCount, offCount] = [0, 0]
    const portCount = Array.from(h.splitters.values()).reduce((total, item) => {
      item.ports.forEach((item) => {
        if (item.status === PortStatus.ON) {
          onCount++
        } else {
          offCount++
        }
      })
      return total + item.ports.size
    }, 0)

    return (
      <>
        <b>{machiRoomName}</b>合计{splitterCount}台分光器，合计{portCount}个端口，其中在用<b>{onCount}</b>
        个端口，空闲端口
        <b>{offCount}</b>
        匹配系统录入数据，其中录入准确XX个端口（第三点的匹配结果），录入有误XX个端口（第三点的匹配结果）。
      </>
    )
  })
}

/**
 * 测试表格-设备端口：
  1、原数据端口的sheet是要核对的，导出表就是从系统上导出来和原数据端口进行核对；
  2、“原数据端口”的C列和“导出表”的A列核对机房是否一致；
  3、“原数据端口”的E&F列和“导出表”的B&C列是匹配条件，“原数据端口”的H列对应“导出表”的E列是匹配结果；
  4、针对原数据端口表输出报告模板如下：

  description:
    XXX机房（C列）合计XX台分光器（看能不能针对E列去重后统计数量，不能就后面人工写入就可以了）
    合计XX个端口（直接统计数据的行数就可以了），
    其中在用XX个端口（G列统计）、空闲端口XX个（G列）；
    匹配系统录入数据，其中录入准确XX个端口（第三点的匹配结果），录入有误XX个端口（第三点的匹配结果）。
 */
function App() {
  const [describtipn, setDescription] = useState<JSX.Element[]>([])

  const analyzeWorkBook: AnalyzeWorkBook = (workBook) => {
    const handleSheet = workBook.Sheets[SHEET_NAME.handle]
    const originSheet = workBook.Sheets[SHEET_NAME.origin]

    const handleData = formatSheetData(handleSheet)
    const originData = formatSheetData(originSheet)
    const description = diffEquipmentData(handleData, originData)
    setDescription(description)
    //

    return {
      handleData,
      originData,
    }
  }

  const handleUploadFile: HandleUploadFile = (w) => {
    // const { handleData, originData } = analyzeWorkBook(w)
    analyzeWorkBook(w)

    console.log('🤔 ~ App ~ w:', w)
  }

  const handleReset = () => {
    setDescription([])
  }

  const handleCheckFile: HandleCheckFile = (w) => {
    const { Sheets } = w
    const handleSheet = Sheets[SHEET_NAME.handle]
    const originSheet = Sheets[SHEET_NAME.origin]

    if (!handleSheet || !originSheet) {
      message.error('上传的数据表数据格式有误，请检查！')
      console.error('必须要有原端口和导出表两个 Sheet!')
      return false
    } else if (!handleSheet['!ref']) {
      message.error('上传的原端口数据表数据异常，请检查')
      console.error('上传了空表')
      return false
    } else if (!originSheet['!ref']) {
      message.error('上传的导出表数据表数据异常，请检查')
      console.error('上传了空表')
      return false
    }
    message.success('上传成功')
    return true
  }

  return (
    <>
      <Flex gap="middle">
        <UploadBtn
          style={{ width: '380px' }}
          onUpload={handleUploadFile}
          onCheck={handleCheckFile}
          onReset={handleReset}
        />
        <Result style={{ width: '100%' }} result={ResultStatus.PERFECT} description={describtipn} />
      </Flex>
    </>
  )
}

export default App

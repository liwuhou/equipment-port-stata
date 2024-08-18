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
  OFF,
  ON,
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

const formatSheetData = (
  s: WorkSheet,
  columns: [string, string, string, string, string],
  initRow = 3
): Map<string, MachinRoom> => {
  const [a, b, c, d, e] = columns
  const map = new Map<string, MachinRoom>()
  const maxLength = parseInt(s['!ref']?.split(':')?.[1]?.replace?.(/[a-z]/gi, '') ?? '0', 10)

  if (maxLength <= initRow) return map

  for (let i = initRow; i <= maxLength; i++) {
    const cell = `${a}${i}`
    const roomName = Reflect.get(s, cell)?.v ?? ''

    if (roomName) {
      if (!map.has(roomName)) {
        map.set(roomName, { name: roomName, splitters: new Map() as MachinRoom['splitters'] })
      }
      const machinRoom = map.get(roomName)!
      const { splitters } = machinRoom

      const splitterName = Reflect.get(s, `${b}${i}`)?.v ?? ''

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

        const portName = Reflect.get(s, `${c}${i}`)?.v ?? ''

        if (portName) {
          if (!ports.has(portName)) {
            const next = Reflect.get(s, `${d}${i}`)?.v?.trim() || ''
            const status = Reflect.get(s, `${e}${i}`)?.v?.trim() === '空闲' ? PortStatus.OFF : PortStatus.ON
            ports.set(portName, { name: portName, belongTo: splitter, status, next })
          }
        }
      }
    }
  }

  return map
}

const countSplitterPorts = (handleSplitters: Map<string, Splitter>, originSplitters: Map<string, Splitter>) => {
  let [on, off, all, right, wrong] = [0, 0, 0, 0, 0]

  for (const [key, handleSplitter] of handleSplitters) {
    all += handleSplitter.ports.size
    const originSplitter = originSplitters.get(key)
    if (handleSplitter.name !== originSplitter?.name) {
      wrong += handleSplitter.ports.size
      continue
    }

    for (const [key, handlePort] of handleSplitter.ports) {
      const originPort = originSplitter.ports.get(key)
      if (handlePort.status === PortStatus.ON) {
        on++
      } else {
        off++
      }
      if (handlePort.name !== originPort?.name) {
        wrong++
        continue
      }
      if (handlePort.status !== originPort?.status) {
        wrong++
        continue
      }
      if (handlePort.next !== originPort?.next) {
        wrong++
        continue
      }
      right++
    }
  }

  return {
    on,
    off,
    all,
    right,
    wrong,
  }
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
  const [description, setDescription] = useState<JSX.Element[]>([])
  const [result, setResult] = useState<ResultStatus>(ResultStatus.PERFECT)

  const analyzeWorkBook: AnalyzeWorkBook = (workBook) => {
    const handleSheet = workBook.Sheets[SHEET_NAME.handle]
    const originSheet = workBook.Sheets[SHEET_NAME.origin]

    const handleData = formatSheetData(handleSheet, ['C', 'E', 'F', 'H', 'G'])
    const originData = formatSheetData(originSheet, ['A', 'B', 'C', 'E', 'N'], 2)
    diffEquipmentData(handleData, originData)

    return {
      handleData,
      originData,
    }
  }

  const diffEquipmentData = (handleData: Map<string, MachinRoom>, originData: Map<string, MachinRoom>): void => {
    let res = []

    for (const [key, handleMachineRoom] of handleData) {
      const originMachineRoom = originData.get(key)
      if (!originMachineRoom) {
        res.push(
          <div>
            <b>{handleMachineRoom.name}</b>录入有误！
          </div>
        )
        break
      }

      const machiRoomName = handleMachineRoom.name
      const splitterCount = handleMachineRoom.splitters.size
      const { on, off, all, right, wrong } = countSplitterPorts(
        handleMachineRoom.splitters,
        originMachineRoom.splitters
      )
      if (right === all) {
        setResult(ResultStatus.PERFECT)
      } else if (wrong > right) {
        setResult(ResultStatus.ERROR)
      } else {
        setResult(ResultStatus.SUCCESS)
      }

      res.push(
        <div>
          <b>{machiRoomName}</b>合计{splitterCount}台分光器，合计<b>{all}</b>个端口，其中<b>{on}</b>
          个在用端口，<b>{off}</b>个空闲端口
          <div>
            匹配系统录入数据，其中录入准确<b>{right}</b>个端口，录入有误<b>{wrong}</b>个端口。
          </div>
        </div>
      )
    }

    setDescription(res)
  }

  const handleUploadFile: HandleUploadFile = (w) => {
    analyzeWorkBook(w)
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
        <UploadBtn onUpload={handleUploadFile} onCheck={handleCheckFile} onReset={handleReset} />
        <Result style={{ width: '100%' }} result={result} description={description} />
      </Flex>
    </>
  )
}

export default App

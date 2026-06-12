import {
  Alert,
  Button,
  Card,
  Col,
  Form,
  Input,
  InputNumber,
  Row,
  Space,
  Statistic,
  Table,
  Tag,
  Typography,
  message,
} from 'antd';
import type { ColumnsType } from 'antd/es/table';
import { BellOutlined, CheckCircleOutlined, ReloadOutlined, SendOutlined } from '@ant-design/icons';
import { useEffect, useMemo, useState } from 'react';

import {
  createSalesLead,
  fetchSalesDashboard,
  matchTransaction,
  type PaymentMatch,
  type PipelineSummary,
  type SalesLead,
} from '../api/sales';

const won = new Intl.NumberFormat('ko-KR');

const statusLabels: Record<string, string> = {
  lead: '선점',
  contract: '계약 완료',
  waiting_payment: '입금 대기',
  paid: '입금 완료',
};

const statusColors: Record<string, string> = {
  lead: 'processing',
  contract: 'blue',
  waiting_payment: 'warning',
  paid: 'success',
};

export function DashboardPage() {
  const [summary, setSummary] = useState<PipelineSummary | null>(null);
  const [leads, setLeads] = useState<SalesLead[]>([]);
  const [matches, setMatches] = useState<PaymentMatch[]>([]);
  const [loading, setLoading] = useState(false);
  const [leadForm] = Form.useForm();
  const [transactionForm] = Form.useForm();

  const loadDashboard = async () => {
    setLoading(true);
    try {
      const data = await fetchSalesDashboard();
      setSummary(data.summary);
      setLeads(data.leads);
      setMatches(data.matches);
    } catch {
      message.error('대시보드 데이터를 불러오지 못했습니다.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void loadDashboard();
  }, []);

  const leadColumns: ColumnsType<SalesLead> = useMemo(
    () => [
      { title: '거래처', dataIndex: 'customer_name', key: 'customer_name' },
      { title: '담당자', dataIndex: 'owner_name', key: 'owner_name', responsive: ['md'] },
      {
        title: '예상 금액',
        dataIndex: 'expected_amount',
        key: 'expected_amount',
        align: 'right',
        render: (value: number) => `${won.format(value)}원`,
      },
      {
        title: '상태',
        dataIndex: 'status',
        key: 'status',
        render: (value: string) => <Tag color={statusColors[value] ?? 'default'}>{statusLabels[value] ?? value}</Tag>,
      },
    ],
    [],
  );

  const matchColumns: ColumnsType<PaymentMatch> = useMemo(
    () => [
      { title: '거래처', dataIndex: 'customer_name', key: 'customer_name' },
      { title: '입금자명', dataIndex: 'depositor_name', key: 'depositor_name', responsive: ['md'] },
      {
        title: '입금액',
        dataIndex: 'amount',
        key: 'amount',
        align: 'right',
        render: (value: number) => `${won.format(value)}원`,
      },
      {
        title: '매칭',
        dataIndex: 'matched_rule',
        key: 'matched_rule',
        render: (value: string) => (value === 'vat_included' ? 'VAT 포함' : '정확 일치'),
      },
    ],
    [],
  );

  const handleCreateLead = async (values: {
    customer_name: string;
    owner_name: string;
    owner_contact?: string;
    meeting_note?: string;
    expected_amount: number;
  }) => {
    try {
      await createSalesLead(values);
      leadForm.resetFields();
      message.success('영업 선점이 등록되었습니다.');
      await loadDashboard();
    } catch {
      message.error('이미 선점된 거래처이거나 입력값을 확인해야 합니다.');
    }
  };

  const handleMatchTransaction = async (values: { depositor_name: string; amount: number }) => {
    try {
      const match = await matchTransaction(values);
      transactionForm.resetFields();
      if (match) {
        message.success(`${match.customer_name} 입금이 자동 매칭되었습니다.`);
      } else {
        message.warning('매칭되는 선점/계약이 없습니다.');
      }
      await loadDashboard();
    } catch {
      message.error('입금 매칭 처리에 실패했습니다.');
    }
  };

  return (
    <Space className="dashboard-page" direction="vertical" size={18}>
      <div className="page-heading">
        <div>
          <Typography.Title level={3}>영업 입금 자동화 대시보드</Typography.Title>
          <Typography.Text type="secondary">선점 등록부터 VAT 포함 입금 매칭까지 한 화면에서 확인합니다.</Typography.Text>
        </div>
        <Button icon={<ReloadOutlined />} onClick={loadDashboard} loading={loading}>
          새로고침
        </Button>
      </div>

      <Row gutter={[16, 16]}>
        <Col xs={24} md={6}>
          <Card>
            <Statistic title="전체 선점" value={summary?.total_leads ?? 0} suffix="건" />
          </Card>
        </Col>
        <Col xs={24} md={6}>
          <Card>
            <Statistic title="입금 대기" value={summary?.waiting_payment ?? 0} suffix="건" />
          </Card>
        </Col>
        <Col xs={24} md={6}>
          <Card>
            <Statistic title="입금 완료" value={summary?.paid ?? 0} suffix="건" prefix={<CheckCircleOutlined />} />
          </Card>
        </Col>
        <Col xs={24} md={6}>
          <Card>
            <Statistic title="확인 입금액" value={summary?.total_paid_amount ?? 0} formatter={(v) => `${won.format(Number(v))}원`} />
          </Card>
        </Col>
      </Row>

      <Row gutter={[16, 16]}>
        <Col xs={24} lg={10}>
          <Card title="모바일 선점 등록">
            <Form form={leadForm} layout="vertical" onFinish={handleCreateLead}>
              <Form.Item name="customer_name" label="거래처명" rules={[{ required: true }]}>
                <Input placeholder="예: OO기업" />
              </Form.Item>
              <Form.Item name="owner_name" label="영업 담당자" rules={[{ required: true }]}>
                <Input placeholder="예: 김영업" />
              </Form.Item>
              <Form.Item name="owner_contact" label="알림 수신처">
                <Input placeholder="전화번호 또는 협업툴 ID" />
              </Form.Item>
              <Form.Item name="expected_amount" label="예상 계약 금액" rules={[{ required: true }]}>
                <InputNumber className="full-width" min={1} step={100000} formatter={(v) => `${won.format(Number(v ?? 0))}`} />
              </Form.Item>
              <Form.Item name="meeting_note" label="미팅 내용">
                <Input.TextArea rows={3} placeholder="현장 미팅 메모" />
              </Form.Item>
              <Button type="primary" htmlType="submit" icon={<SendOutlined />} block>
                선점 등록
              </Button>
            </Form>
          </Card>
        </Col>

        <Col xs={24} lg={14}>
          <Card title="입금 감지 테스트">
            <Alert
              showIcon
              type="info"
              icon={<BellOutlined />}
              message="금융 스크래핑 연동 전까지는 수동 거래 입력으로 자동 매칭 로직을 검증합니다."
            />
            <Form className="transaction-form" form={transactionForm} layout="inline" onFinish={handleMatchTransaction}>
              <Form.Item name="depositor_name" rules={[{ required: true }]}>
                <Input placeholder="입금자명" />
              </Form.Item>
              <Form.Item name="amount" rules={[{ required: true }]}>
                <InputNumber min={1} step={100000} placeholder="입금액" formatter={(v) => `${won.format(Number(v ?? 0))}`} />
              </Form.Item>
              <Button type="primary" htmlType="submit" icon={<CheckCircleOutlined />}>
                매칭 실행
              </Button>
            </Form>
            <Table
              className="section-table"
              rowKey="id"
              columns={matchColumns}
              dataSource={matches}
              loading={loading}
              pagination={{ pageSize: 5 }}
            />
          </Card>
        </Col>
      </Row>

      <Row gutter={[16, 16]}>
        <Col xs={24} lg={16}>
          <Card title="영업 파이프라인">
            <Table rowKey="id" columns={leadColumns} dataSource={leads} loading={loading} pagination={{ pageSize: 8 }} />
          </Card>
        </Col>
        <Col xs={24} lg={8}>
          <Card title="미수금 리스크">
            <Space direction="vertical" className="full-width">
              {(summary?.overdue_risk ?? []).map((lead) => (
                <div className="risk-item" key={lead.id}>
                  <strong>{lead.customer_name}</strong>
                  <span>{won.format(lead.expected_amount)}원</span>
                </div>
              ))}
              {summary?.overdue_risk.length === 0 && <Typography.Text type="secondary">대기 중인 미수 리스크가 없습니다.</Typography.Text>}
            </Space>
          </Card>
        </Col>
      </Row>
    </Space>
  );
}

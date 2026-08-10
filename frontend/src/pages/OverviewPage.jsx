import { Link } from 'react-router-dom';

const OBJECTIVES = [
  {
    to: '/sales-order',
    theme: 'sales',
    label: 'Sales Order',
    title: 'Generate ESK Payment File',
    desc: 'Build the sales order and payment listing validation file from SK and NC sources.',
  },
  {
    to: '/stripe',
    theme: 'stripe',
    label: 'Stripe',
    title: 'Generate Stripe Advice File',
    desc: 'Combine Stripe payout data with ESK payment details into Stripe advice.',
  },
  {
    to: '/gpay',
    theme: 'gpay',
    label: 'GPay',
    title: 'Generate GPay Advice File',
    desc: 'Create GPay payment advice from Grab transactions and ESK listing.',
  },
  {
    to: '/shopback',
    theme: 'shopback',
    label: 'ShopBack',
    title: 'Generate Shopback Advice File',
    desc: 'Create ShopBack payment advice from Shopback transactions and ESK listing.',
  },
];

export default function OverviewPage() {
  return (
    <div className="dashboard">
      <header className="page-hero">
        <h2 className="page-hero__title">Overview</h2>
      </header>

      <section className="overview-objectives">
        <h2>Objective</h2>
        <ul>
          <li>Generate ESK Payment File</li>
          <li>Generate Stripe Advice File</li>
          <li>Generate GPay Advice File</li>
          <li>Generate Shopback Advice File</li>
        </ul>
      </section>

      <div className="overview-grid">
        {OBJECTIVES.map((item) => (
          <Link
            key={item.to}
            to={item.to}
            className={`overview-card overview-card--${item.theme}`}
          >
            <span className="overview-card__label">{item.label}</span>
            <h3 className="overview-card__title">{item.title}</h3>
            <p className="overview-card__desc">{item.desc}</p>
          </Link>
        ))}
      </div>
    </div>
  );
}

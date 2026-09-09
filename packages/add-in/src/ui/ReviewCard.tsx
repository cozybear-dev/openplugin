import { Body1, Button, Caption1 } from "@fluentui/react-components";

export function ReviewCard(props: {
  items: Array<{ title: string; detail: string }>;
  onApply: () => void;
  onReject: () => void;
  onAlways: () => void;
}) {
  return (
    <section className="op-review">
      <Caption1 className="op-review-kicker">Review changes</Caption1>
      <ul>
        {props.items.map((item, i) => (
          <li key={i}>
            <Body1>{item.title}</Body1>
            {item.detail ? <Caption1>{item.detail}</Caption1> : null}
          </li>
        ))}
      </ul>
      <div className="op-review-actions">
        <Button appearance="primary" onClick={props.onApply}>
          Apply
        </Button>
        <Button appearance="subtle" onClick={props.onReject}>
          Reject
        </Button>
        <Button appearance="transparent" onClick={props.onAlways}>
          Always apply this session
        </Button>
      </div>
    </section>
  );
}

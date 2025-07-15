import * as React from 'react';
import { useState } from 'react';
import styles from './components/Tempalte120.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IAccordionSection {
  title: string;
  content: JSX.Element | string;
}

export interface IAccordionTemplateProps {
  sections: IAccordionSection[];
}

const AccordionTemplate: React.FC<IAccordionTemplateProps> = ({ sections }) => {
  const [openIndexes, setOpenIndexes] = useState<number[]>([0]);

  const toggle = (idx: number) => {
    setOpenIndexes(openIndexes.includes(idx)
      ? openIndexes.filter(i => i !== idx)
      : [...openIndexes, idx]);
  };

  return (
    <div className={styles.accordion}>
      {sections.map((section, idx) => (
        <div key={idx} className={styles.accordionSection}>
          <div className={styles.accordionHeader} onClick={() => toggle(idx)}>
            <span className={styles.accordionTitle}>{section.title}</span>
            <Icon iconName={openIndexes.includes(idx) ? "ChevronDown" : "ChevronRight"} className={styles.accordionIcon} />
          </div>
          {openIndexes.includes(idx) && (
            <div className={styles.accordionContent}>{section.content}</div>
          )}
        </div>
      ))}
    </div>
  );
};

export default AccordionTemplate;
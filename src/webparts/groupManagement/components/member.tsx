import { Persona, PersonaSize } from "@fluentui/react";
import * as React from "react";
// import styles from "./GroupManagement.module.scss";

interface IMemberProps {
  name: string;
  email: string;
}

const Member: React.FC<IMemberProps> = ({ name, email }) => {
  return (
    <div>
      <Persona
        text={name}
        secondaryText={email}
        size={PersonaSize.size32}
        showSecondaryText={true}
      />
    </div>
  );
};

export default Member;

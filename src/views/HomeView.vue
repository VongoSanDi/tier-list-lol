<template>
  <div class="container">
    <div class="top-section">
      <input type="file" accept=".xls, .xlsx" @change="importExcel" />
      <button @click="test">test</button>
      <div v-if="sheets.length > 0" class="mt-4">
        <label for="sheet-select">Choisi une page du tableur</label>
        <select id="sheet-select" v-model="selectedSheet"
          @change="e => handleSheetChange((e.target as HTMLSelectElement).value)">
          <option v-for="sheet in sheets" :key="sheet.name" :value="sheet.name">
            {{ sheet.name }}
          </option>
        </select>
      </div>
      <button @click="save" class="save">Sauvegarder</button>
    </div>
    <div class="middle-section">
      <!-- joueurs classés -->
      <div class="waiting-section">
        <div v-for="(players, position, indexColumn) in rankedPlayers" :key="indexColumn" class="position-column">
          <h3 class="position-header">{{ position }}</h3>
          <div class="position-players">
            <div v-for="(player, index) in players" :key="index" class="player" @dragover.prevent :draggable="!!player"
              @dragstart="player && dragStart($event, player, position, index)" @drop="drop($event, position, index)">
              <div v-if="player">
                {{ player.name }}
              </div>
              <div v-else class="empty-slot">
                Emplacement vide
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- Joueurs en attente -->
      <span>Joueur en attente</span>
      <div class="waiting-section">
        <div v-for="(players, position, indexColumn) in waitingPlayers" :key="indexColumn" class="position-column">
          <div class="position-header">{{ position }}</div>
          <div class="position-players">
            <div v-for="(player, index) in players" :key="player.id" class="player" draggable="true"
              @dragstart="dragStart($event, player, position, index)">
              {{ player.name }}
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from "vue"
import { read, utils, writeFileXLSX } from 'xlsx'
import type { WorkBook } from "xlsx"
import { v4 as uuidv4 } from 'uuid';

type Role = "top" | "jungle" | "mid" | "adc" | "support";

interface Player {
  id: string;
  name: string;
}

const sheets = ref<{ name: string }[]>([])
const selectedSheet = ref('')
const workbook = ref<WorkBook | null>(null)

const topPlayers = ref<Player[]>([])
const junglePlayers = ref<Player[]>([])
const midPlayers = ref<Player[]>([])
const adcPlayers = ref<Player[]>([])
const supportPlayers = ref<Player[]>([])

const rankedPlayers = ref<Record<Role, (Player | null)[]>>({
  top: Array.from({ length: 10 }, () => null),
  jungle: Array.from({ length: 10 }, () => null),
  mid: Array.from({ length: 10 }, () => null),
  adc: Array.from({ length: 10 }, () => null),
  support: Array.from({ length: 10 }, () => null)
});

const waitingPlayers = computed<Record<Role, Player[]>>(() => ({
  top: topPlayers.value,
  jungle: junglePlayers.value,
  mid: midPlayers.value,
  adc: adcPlayers.value,
  support: supportPlayers.value,
}))


const importExcel = async (event: Event) => {
  const target = event.target as HTMLInputElement
  if (!target.files?.length) return

  const file = target.files[0]

  const buffer = await file.arrayBuffer()
  workbook.value = read(new Uint8Array(buffer), { type: 'array' })

  sheets.value = workbook.value.SheetNames.map(name => ({ name }))

  //Charge la première feuille par défaut s'il y en a qu'une seule
  if (workbook.value.SheetNames.length === 1) {
    handleSheetChange(workbook.value.SheetNames[0])
  }
}

// Obtenir la donnée contenu dans une cellule
const readCell = (sheetName: string, cellAddress: string) => {
  if (!workbook.value) return null

  const worksheet = workbook.value.Sheets[sheetName]
  if (!worksheet) return null

  const cell = worksheet[cellAddress]
  return cell ? cell.v : null // .v contient la valeur de la cellule
}

// Obtenir les données contenues dans une plage de cellules
const readRange = (sheetName: string, startCell: string, endCell: string) => {
  if (!workbook.value) return []

  const worksheet = workbook.value.Sheets[sheetName]
  if (!worksheet) return []

  const range = utils.decode_range(`${startCell}:${endCell}`)
  const values = []

  for (let row = range.s.r; row <= range.e.r; row++) {
    const rowValues = []
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = utils.encode_cell({ r: row, c: col })
      const value = readCell(sheetName, cellAddress)
      rowValues.push(value)
    }
    values.push(rowValues)
  }

  return values
}

const handleSheetChange = (sheetName: string) => {
  if (!workbook.value) return

  selectedSheet.value = sheetName

  topPlayers.value = generatePlayers('B20');
  junglePlayers.value = generatePlayers('C20');
  midPlayers.value = generatePlayers('D20');
  adcPlayers.value = generatePlayers('E20');
  supportPlayers.value = generatePlayers('F20');
}

const generatePlayers = (startRange: string) => {
  return readRange(selectedSheet.value, startRange, startRange.replace(/\d+/, '24'))
    .flatMap(p => p)
    .filter(name => name) // Pour éviter les valeurs nulles ou vides
    .map(name => ({ id: uuidv4(), name }));
};

const test = () => {
  console.log('test1', waitingPlayers.value)
  console.log('test2', rankedPlayers)
}

const save = () => {
  console.log('save')
  if (!workbook.value || !selectedSheet.value) return;
  const worksheet = workbook.value.Sheets[selectedSheet.value];
  if (!worksheet) return;

  // Définir où enregistrer les joueurs (Ex: rang 30 à 34)
  const positions: Record<Role, string> = {
    top: "B2",
    jungle: "E2",
    mid: "H2",
    adc: "K2",
    support: "N2",
  };

  // Mettre à jour les cellules avec les joueurs classés
  (Object.keys(rankedPlayers.value) as Role[]).forEach((role) => {
    rankedPlayers.value[role].forEach((player, index) => {
      console.log('save.index', index)
      console.log('save.player', player)
      if (player) {
        const cellAddress = utils.encode_cell({
          r: 1 + index, // Ligne de départ (30) + index
          c: utils.decode_col(positions[role][0]), // Colonne correspondante
        });

        // Écriture de la valeur dans la cellule
        worksheet[cellAddress] = { v: player.name, t: "s" };
      }
    });
  });
  console.log("✅ Modifications appliquées au fichier XLSX !");

  writeFileXLSX(workbook.value, "updated_file.xlsx", { bookType: "xlsx", type: "file" });
  console.log("✅ Fichier XLSX mis à jour !");
}

const dragStart = (event: DragEvent, player: Player, positionStart: Role, indexStart: number) => {
  event.dataTransfer?.setData("text/plain", JSON.stringify({ playerId: player.id, positionStart, indexStart }));
};

const drop = (event: DragEvent, positionDrop: Role, indexDrop: number) => {

  event.preventDefault();
  const data = event.dataTransfer?.getData("text/plain");
  if (!data) return;
  console.log('drop.data', data)
  console.log('drop.indexDrop', indexDrop)

  const { playerId, positionStart, indexStart } = JSON.parse(data) as {
    playerId: string;
    positionStart: Role;
    indexStart: number;
  };

  // Vérifie si le joueur est déplacé dans la même position
  if (positionStart === positionDrop && indexStart === indexDrop) {
    console.log("Le joueur est déjà à cet endroit.");
    return;
  }

  const sourceList = waitingPlayers.value[positionStart] as Player[];
  const targetList = rankedPlayers.value[positionDrop] as Player[]

  // Supprime le joueur de sa position d'origine
  const player = sourceList.find((p) => p.id === playerId);
  if (!player) return;
  sourceList.splice(indexStart, 1);

  // Récupère la liste cible
  console.log('drop', targetList)
  console.log('drop2', targetList[indexDrop])
  console.log('drop3', rankedPlayers.value[positionDrop])
  console.log("Type de targetList[indexDrop]:", typeof targetList[indexDrop]);
  console.log("Valeur réelle :", JSON.stringify(targetList[indexDrop]));

  // Vérifie si la position cible est déjà occupée
  if (targetList[indexDrop]) {
    console.log(`L'emplacement à l'index ${indexDrop} est déjà occupé.`);

    // Déplace les joueurs en décalant vers le bas
    targetList.splice(indexDrop, 0, player);  // Décale et insère
  } else {
    // Insère directement si l'emplacement est vide
    targetList.splice(indexDrop, 1, player);
  }
};

</script>
